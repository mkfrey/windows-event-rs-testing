use std::ffi::c_void;
use std::fmt::Debug;
use std::ptr::{null, null_mut};
use std::slice::from_raw_parts;

use crate::WindowsConversionTo;
use crate::WindowsError;

use windows_result::HRESULT;
use windows_strings::HSTRING;
use windows_sys::core::{GUID, PCSTR as PCSTR_SYS, PCWSTR as PCWSTR_SYS};
use windows_sys::Win32::Foundation::{ERROR_INSUFFICIENT_BUFFER, FILETIME, SYSTEMTIME};
use windows_sys::Win32::Security::{PSID, SID};
use windows_sys::Win32::System::EventLog::*;

static ZERO_BUFFER_SIZE: u32 = 0;
static NULL_EVT_HANDLE: EVT_HANDLE = 0 as EVT_HANDLE;

/// Windows event handle wrapper providing additional functionality
/// IMPORTANT: The wrapped handle needs to be valid for the entire lifetime of the struct.
pub struct WindowsEvent<'a> {
    handle: &'a EVT_HANDLE,
}

impl<'a> WindowsEvent<'a> {
    pub fn new(handle: &'a EVT_HANDLE) -> Self {
        Self { handle }
    }

    fn render_generic(
        &self,
        valuepaths: &[PCWSTR_SYS],
        context_flags: u32,
        render_flags: u32,
    ) -> Result<(Vec<u8>, u32), String> {
        let mut buffer_used: u32 = 0;
        let mut property_count: u32 = 0;

        println!("{:?}", valuepaths);
        println!("{:?}", valuepaths.as_ptr());

        let render_context =
            if render_flags == EvtRenderEventXml || render_flags == EvtRenderBookmark {
                // For rendering XML or bookmarks, context has to be NULL
                EventRenderContext::create_null()
            } else {
                match EventRenderContext::create(
                    valuepaths.len() as u32,
                    valuepaths.as_ptr(),
                    context_flags,
                ) {
                    Ok(context) => context,
                    Err(error) => {
                        return Err(format!(
                            "Error trying to create render context: {:?}",
                            error.message()
                        ));
                    }
                }
            };

        // Render the event values with zero length buffer to determine size.
        // Hint: EvtRender may return 1 if the property count of the event is 0.
        unsafe {
            EvtRender(
                render_context.as_ptr(),
                *self.handle,
                render_flags,
                ZERO_BUFFER_SIZE,
                null_mut(),
                &mut buffer_used,
                &mut property_count,
            )
        };

        let last_error = WindowsError::from_win32();

        // ... and to receive the error ERROR_INSUFFICIENT_BUFFER, if anything needs to be rendered.
        if last_error.code().is_err() {
            if last_error.code() != HRESULT::from_win32(ERROR_INSUFFICIENT_BUFFER) {
                return Err(format!(
                    "Error trying to determine buffer size: {:?}",
                    last_error.code()
                ));
            }
        }

        let mut buffer: Vec<u8> = vec![0; buffer_used as usize];

        unsafe {
            EvtRender(
                render_context.as_ptr(),
                *self.handle,
                render_flags,
                buffer.len() as u32,
                buffer.as_mut_ptr() as *mut c_void,
                &mut buffer_used,
                &mut property_count,
            )
        };

        let last_error = WindowsError::from_win32();

        if last_error.code().is_err() {
            return Err(format!(
                "Error when trying to render event: {:?}",
                last_error.message()
            ));
        }

        return Ok((buffer, property_count));
    }

    pub fn render_system_context(&self) -> Result<EventSystemContext, String> {
        let (raw_buffer, property_count) =
            self.render_generic(&[], EvtRenderContextSystem, EvtRenderEventValues)?;
        let buffer = unsafe { EventVariantBuffer::from_raw_buffer(raw_buffer, property_count) };
        Ok(unsafe { EventSystemContext::from_variant_buffer(&buffer) })
    }

    pub fn render_user_context(&self) -> Result<Vec<Result<EventVariantValue, String>>, String> {
        let (raw_buffer, property_count) =
            self.render_generic(&[], EvtRenderContextUser, EvtRenderEventValues)?;
        let buffer = unsafe { EventVariantBuffer::from_raw_buffer(raw_buffer, property_count) };
        Ok(buffer.into_iter().collect())
    }

    pub fn render_xml(&self) -> Result<String, String> {
        let (raw_buffer, _) = self.render_generic(&[], 0, EvtRenderEventXml)?;

        Ok((raw_buffer.as_ptr() as *const u16).win_into())
    }

    pub fn render_description(&self) -> Result<String, String> {
        let pathspec_system_provider = HSTRING::from("Event/System/Provider/@Name");
        let pathspec_rendering_inf = HSTRING::from("Event/RenderingInfo/Message");

        let pathspecs = [
            pathspec_system_provider.as_ptr(),
            pathspec_rendering_inf.as_ptr(),
        ];

        let (raw_buffer, property_count) = self.render_generic(
            pathspecs.as_slice(),
            EvtRenderContextValues,
            EvtRenderEventValues,
        )?;

        let buffer = unsafe { EventVariantBuffer::from_raw_buffer(raw_buffer, property_count) };

        match buffer.get_property_value(0) {
            // If the event was not forwarded, Null will be returned, otherwise a String.
            EventVariantValue::Null() => {
                todo!()
            }
            EventVariantValue::String(str) => {
                str
            }
        }
    }
}

/// Convenience wrapper around a buffer containing `EVENT_VARIANT` objects.
#[derive(Debug)]
pub struct EventVariantBuffer {
    buffer: Vec<u8>,
    property_count: u32,
}

impl EventVariantBuffer {
    /// Create a new wrapper from a raw byte buffer and the number of properties it contains.
    ///
    /// Requires the buffer to contain `EVENT_VARIANT` objects as returned by `EvtRender`.
    /// No checks are performed.
    pub unsafe fn from_raw_buffer(buffer: Vec<u8>, property_count: u32) -> Self {
        Self {
            buffer,
            property_count,
        }
    }

    pub fn property_count(&self) -> u32 {
        self.property_count
    }

    pub fn get_property_value(&self, index: u32) -> Option<Result<EventVariantValue, String>> {
        if index >= self.property_count {
            return None;
        }

        let raw_variant = unsafe { self.index(index as isize) };

        Some((*raw_variant).try_into())
    }

    pub fn as_ptr(&self) -> *const EVT_VARIANT {
        self.buffer.as_ptr() as *const EVT_VARIANT
    }

    pub fn buffer_size(&self) -> usize {
        self.buffer.len()
    }

    pub unsafe fn index(&self, offset: isize) -> &EVT_VARIANT {
        unsafe { self.as_ptr().offset(offset).as_ref().unwrap_unchecked() }
    }
}

impl<'a> IntoIterator for &'a EventVariantBuffer {
    type Item = Result<EventVariantValue, String>;

    type IntoIter = EventVariantBufferIterator<'a>;

    fn into_iter(self) -> Self::IntoIter {
        EventVariantBufferIterator {
            buffer: self,
            index: 0,
        }
    }
}

pub struct EventVariantBufferIterator<'a> {
    buffer: &'a EventVariantBuffer,
    index: u32,
}

impl<'a> Iterator for EventVariantBufferIterator<'a> {
    type Item = Result<EventVariantValue, String>;

    fn next(&mut self) -> Option<Self::Item> {
        let value = self.buffer.get_property_value(self.index);
        self.index += 1;
        value
    }
}

/// Rust representation of a EVT_VARIANT, owning all of its data.
///
/// This enum and its implementation assumes that any
/// data field of a variant referenced by pointer is not a nullpointer.
///
/// This is not guaranteed if the fields of rendering an event
/// are fixed, e.g. when rendering a system event.
///
/// Bigger data fields are intentionally boxed. No guarantees are provided
/// regarding the validity of the EvtHandle value.
///
pub enum EventVariantValue {
    Null,
    Bool(bool),
    SByte(i8),
    Int16(i16),
    Int32(i32),
    Int64(i64),
    Byte(u8),
    UInt16(u16),
    UInt32(u32),
    UInt64(u64),
    Single(f32),
    Double(f64),
    FileTime(FILETIME),
    SysTime(Box<SYSTEMTIME>),
    Guid(Box<GUID>),
    HexInt32(u32),
    HexInt64(u64),
    String(String),
    AnsiString(String),
    Binary(Vec<u8>),
    Sid(Box<SID>),
    SizeT(usize),
    BoolArr(Vec<bool>),
    SByteArr(Vec<i8>),
    Int16Arr(Vec<i16>),
    Int32Arr(Vec<i32>),
    Int64Arr(Vec<i64>),
    ByteArr(Vec<u8>),
    UInt16Arr(Vec<u16>),
    UInt32Arr(Vec<u32>),
    UInt64Arr(Vec<u64>),
    SingleArr(Vec<f32>),
    DoubleArr(Vec<f64>),
    FileTimeArr(Vec<FILETIME>),
    SysTimeArr(Vec<SYSTEMTIME>),
    GuidArr(Vec<GUID>),
    HexInt32Arr(Vec<u32>),
    HexInt64Arr(Vec<u64>),
    StringArr(Vec<String>),
    AnsiStringArr(Vec<String>),
    SidArr(Vec<SID>),
    SizeTArr(Vec<usize>),
    EvtHandle(EVT_HANDLE),
    Xml(String),
    XmlArr(Vec<String>),
}

/*impl ToString for EventVariantValue {
    fn to_string(&self) -> String {
        match self {
            EventVariantValue::Null() => "Null".to_owned(),
            EventVariantValue::Bool(val) => {
                if *val {"Bool(true)"} else {"Bool(false)"}
            }.to_owned(),
            EventVariantValue::SByte(val) => format!("")


        }
    }
}*/

impl TryFrom<EVT_VARIANT> for EventVariantValue {
    type Error = String;

    fn try_from(value: EVT_VARIANT) -> Result<Self, Self::Error> {
        let is_array = (value.Type & EVT_VARIANT_TYPE_ARRAY) != 0;
        let value_type = (value.Type & EVT_VARIANT_TYPE_MASK) as i32;
        let count = value.Count as usize;

        if is_array {
            unsafe {
                match value_type {
                    EvtVarTypeString => Ok(Self::StringArr(
                        from_raw_parts(value.Anonymous.StringArr as *const PCWSTR_SYS, count)
                            .iter()
                            .map(|s| (*s).win_into())
                            .collect(),
                    )),
                    EvtVarTypeAnsiString => Ok(Self::AnsiStringArr(
                        from_raw_parts(value.Anonymous.AnsiStringArr as *const PCSTR_SYS, count)
                            .iter()
                            .map(|s| (*s).win_into())
                            .collect(),
                    )),
                    EvtVarTypeSByte => Ok(Self::SByteArr(
                        from_raw_parts(value.Anonymous.SByteArr, count).to_vec(),
                    )),
                    EvtVarTypeByte => Ok(Self::ByteArr(
                        from_raw_parts(value.Anonymous.ByteArr, count).to_vec(),
                    )),
                    EvtVarTypeInt16 => Ok(Self::Int16Arr(
                        from_raw_parts(value.Anonymous.Int16Arr, count).to_vec(),
                    )),
                    EvtVarTypeUInt16 => Ok(Self::UInt16Arr(
                        from_raw_parts(value.Anonymous.UInt16Arr, count).to_vec(),
                    )),
                    EvtVarTypeInt32 => Ok(Self::Int32Arr(
                        from_raw_parts(value.Anonymous.Int32Arr, count).to_vec(),
                    )),
                    EvtVarTypeUInt32 => Ok(Self::UInt32Arr(
                        from_raw_parts(value.Anonymous.UInt32Arr, count).to_vec(),
                    )),
                    EvtVarTypeInt64 => Ok(Self::Int64Arr(
                        from_raw_parts(value.Anonymous.Int64Arr, count).to_vec(),
                    )),
                    EvtVarTypeUInt64 => Ok(Self::UInt64Arr(
                        from_raw_parts(value.Anonymous.UInt64Arr, count).to_vec(),
                    )),
                    EvtVarTypeSingle => Ok(Self::SingleArr(
                        from_raw_parts(value.Anonymous.SingleArr, count).to_vec(),
                    )),
                    EvtVarTypeDouble => Ok(Self::DoubleArr(
                        from_raw_parts(value.Anonymous.DoubleArr, count).to_vec(),
                    )),
                    EvtVarTypeBoolean => Ok(Self::BoolArr(
                        from_raw_parts(value.Anonymous.BooleanArr, count)
                            .iter()
                            .map(|b| *b != 0)
                            .collect(),
                    )),
                    EvtVarTypeGuid => Ok(Self::GuidArr(
                        from_raw_parts(value.Anonymous.GuidArr, count).to_vec(),
                    )),
                    EvtVarTypeSizeT => Ok(Self::SizeTArr(
                        from_raw_parts(value.Anonymous.SizeTArr, count).to_vec(),
                    )),
                    EvtVarTypeFileTime => Ok(Self::FileTimeArr(
                        from_raw_parts(value.Anonymous.FileTimeArr, count).to_vec(),
                    )),
                    EvtVarTypeSysTime => Ok(Self::SysTimeArr(
                        from_raw_parts(value.Anonymous.SysTimeArr, count).to_vec(),
                    )),
                    EvtVarTypeSid => Ok(Self::SidArr(
                        from_raw_parts(value.Anonymous.SidArr as *const *const SID, count)
                            .iter()
                            .map(|s| **s)
                            .collect(),
                    )),
                    EvtVarTypeHexInt32 => Ok(Self::HexInt32Arr(
                        from_raw_parts(value.Anonymous.UInt32Arr, count).to_vec(),
                    )),
                    EvtVarTypeHexInt64 => Ok(Self::HexInt64Arr(
                        from_raw_parts(value.Anonymous.UInt64Arr, count).to_vec(),
                    )),
                    EvtVarTypeEvtXml => Ok(Self::XmlArr(
                        from_raw_parts(value.Anonymous.XmlValArr, count)
                            .iter()
                            .map(|s| (*s as PCWSTR_SYS).win_into())
                            .collect(),
                    )),
                    _ => Err(format!("No mapping found for array type {}", value_type)),
                }
            }
        } else {
            unsafe {
                match value_type {
                    EvtVarTypeNull => Ok(Self::Null),
                    EvtVarTypeString => Ok(Self::String(value.Anonymous.StringVal.win_into())),
                    EvtVarTypeAnsiString => {
                        Ok(Self::AnsiString(value.Anonymous.AnsiStringVal.win_into()))
                    }
                    EvtVarTypeSByte => Ok(Self::SByte(value.Anonymous.SByteVal)),
                    EvtVarTypeByte => Ok(Self::Byte(value.Anonymous.ByteVal)),
                    EvtVarTypeInt16 => Ok(Self::Int16(value.Anonymous.Int16Val)),
                    EvtVarTypeUInt16 => Ok(Self::UInt16(value.Anonymous.UInt16Val)),
                    EvtVarTypeInt32 => Ok(Self::Int32(value.Anonymous.Int32Val)),
                    EvtVarTypeUInt32 => Ok(Self::UInt32(value.Anonymous.UInt32Val)),
                    EvtVarTypeInt64 => Ok(Self::Int64(value.Anonymous.Int64Val)),
                    EvtVarTypeUInt64 => Ok(Self::UInt64(value.Anonymous.UInt64Val)),
                    EvtVarTypeSingle => Ok(Self::Single(value.Anonymous.SingleVal)),
                    EvtVarTypeDouble => Ok(Self::Double(value.Anonymous.DoubleVal)),
                    EvtVarTypeBoolean => Ok(Self::Bool(value.Anonymous.BooleanVal != 0)),
                    EvtVarTypeBinary => Ok(Self::Binary(
                        from_raw_parts(value.Anonymous.BinaryVal, count).to_vec(),
                    )),
                    EvtVarTypeGuid => Ok(Self::Guid(Box::new(*value.Anonymous.GuidVal))),
                    EvtVarTypeSizeT => Ok(Self::SizeT(value.Anonymous.SizeTVal)),
                    EvtVarTypeFileTime => {
                        Ok(Self::FileTime(value.Anonymous.FileTimeVal.win_into()))
                    }
                    EvtVarTypeSysTime => Ok(Self::SysTime(Box::new(*value.Anonymous.SysTimeVal))),
                    EvtVarTypeSid => {
                        Ok(Self::Sid(Box::new(*(value.Anonymous.SidVal as *const SID))))
                    }
                    EvtVarTypeHexInt32 => Ok(Self::HexInt32(value.Anonymous.UInt32Val)),
                    EvtVarTypeHexInt64 => Ok(Self::HexInt64(value.Anonymous.UInt64Val)),
                    EvtVarTypeEvtHandle => Ok(Self::EvtHandle(value.Anonymous.EvtHandleVal)),
                    EvtVarTypeEvtXml => Ok(Self::Xml(value.Anonymous.XmlVal.win_into())),
                    _ => Err(format!("No mapping found for value type {}", value_type)),
                }
            }
        }
    }
}

/// Rust representation of an event render context
pub struct EventRenderContext {
    render_context: EVT_HANDLE,
}

/// Implement trait `Drop` to enforce proper disposal of the underlying windows object.
impl Drop for EventRenderContext {
    fn drop(&mut self) {
        if self.render_context != NULL_EVT_HANDLE {
            unsafe {
                EvtClose(self.render_context);
            };
        }
    }
}

impl EventRenderContext {
    /// Create a new render context with the provided parameters.
    ///
    /// # Parameters
    /// - `valuepathscount`: The number of elements in the `valuepaths` array.
    /// - `valuepaths`: A pointer to an array of strings that specify the names of the values to be rendered.
    /// - `flags`: Flags that specify which context is created.
    pub fn create(
        valuepathscount: u32,
        valuepaths: *const windows_sys::core::PCWSTR,
        flags: u32,
    ) -> Result<Self, WindowsError> {
        let context = unsafe { EvtCreateRenderContext(valuepathscount, valuepaths, flags) };

        if context == NULL_EVT_HANDLE {
            Err(WindowsError::from_win32())
        } else {
            Ok(Self {
                render_context: context,
            })
        }
    }

    /// Create a render context with a `NULL` handle.
    ///
    /// Useful if render call requires the context parameter to be NULL.
    pub fn create_null() -> Self {
        Self {
            render_context: NULL_EVT_HANDLE,
        }
    }

    /// Get the underlying `EVT_HANDLE` of the context.
    pub fn as_ptr(&self) -> EVT_HANDLE {
        self.render_context
    }
}

/// Rust representation of a rendered system context.
///
/// See https://learn.microsoft.com/en-us/windows/win32/api/winevt/ne-winevt-evt_system_property_id for system context values
pub struct EventSystemContext {
    pub provider_name: String,
    pub provider_guid: Option<GUID>,
    pub event_id: u16,
    pub qualifiers: u16,
    pub level: u8,
    pub task: u16,
    pub opcode: u8,
    pub keywords: i64,
    pub time_created: u64,
    pub event_record_id: u64,
    pub activity_id: Option<GUID>,
    pub related_activity_id: Option<GUID>,
    pub process_id: u32,
    pub thread_id: u32,
    pub channel: String,
    pub computer: String,
    pub user_id: Option<SID>,
    pub version: u8,
}

impl EventSystemContext {
    /// Extract the system context data from a variant buffer.
    ///
    /// # Safety
    ///
    /// This function is unsafe because it assumes that the `variant` parameter contains valid system context data.
    /// No checks are performed to ensure the validity of the data, and dereferencing raw pointers is inherently unsafe.
    pub unsafe fn from_variant_buffer(variant: &EventVariantBuffer) -> Self {
        unsafe {
            Self {
                provider_name: variant
                    .index(EvtSystemProviderName as isize)
                    .Anonymous
                    .StringVal
                    .win_into(),
                provider_guid: variant
                    .index(EvtSystemProviderGuid as isize)
                    .Anonymous
                    .GuidVal
                    .as_ref()
                    .cloned(),
                event_id: variant.index(EvtSystemEventID as isize).Anonymous.UInt16Val,
                qualifiers: variant
                    .index(EvtSystemQualifiers as isize)
                    .Anonymous
                    .UInt16Val,
                level: variant.index(EvtSystemLevel as isize).Anonymous.ByteVal,
                task: variant.index(EvtSystemTask as isize).Anonymous.UInt16Val,
                opcode: variant.index(EvtSystemOpcode as isize).Anonymous.ByteVal,
                keywords: variant.index(EvtSystemKeywords as isize).Anonymous.Int64Val,
                time_created: variant
                    .index(EvtSystemTimeCreated as isize)
                    .Anonymous
                    .UInt64Val,
                event_record_id: variant
                    .index(EvtSystemEventRecordId as isize)
                    .Anonymous
                    .UInt64Val,
                activity_id: variant
                    .index(EvtSystemActivityID as isize)
                    .Anonymous
                    .GuidVal
                    .as_ref()
                    .cloned(),
                related_activity_id: variant
                    .index(EvtSystemRelatedActivityID as isize)
                    .Anonymous
                    .GuidVal
                    .as_ref()
                    .cloned(),
                process_id: variant
                    .index(EvtSystemProcessID as isize)
                    .Anonymous
                    .UInt32Val,
                thread_id: variant
                    .index(EvtSystemThreadID as isize)
                    .Anonymous
                    .UInt32Val,
                channel: variant
                    .index(EvtSystemChannel as isize)
                    .Anonymous
                    .StringVal
                    .win_into(),
                computer: variant
                    .index(EvtSystemComputer as isize)
                    .Anonymous
                    .StringVal
                    .win_into(),
                user_id: (variant.index(EvtSystemUserID as isize).Anonymous.SidVal as *const SID)
                    .as_ref()
                    .cloned(),
                version: variant.index(EvtSystemVersion as isize).Anonymous.ByteVal,
            }
        }
    }
}
