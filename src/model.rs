use std::ffi::c_void;
use std::fmt;
use std::fmt::Debug;
use std::ptr::{null, null_mut};
use std::slice::from_raw_parts;

use crate::WindowsConversionTo;
use crate::WindowsError;

use chrono::{DateTime, NaiveDateTime, Utc};

use windows_result::HRESULT;
use windows_strings::HSTRING;
use windows_sys::core::{GUID, PCSTR as PCSTR_SYS, PCWSTR as PCWSTR_SYS};
use windows_sys::Win32::Foundation::ERROR_INSUFFICIENT_BUFFER;
use windows_sys::Win32::Security::SID;
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
                    if valuepaths.len() > 1 {
                        valuepaths.as_ptr()
                    } else {
                        null()
                    },
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

        // TODO: Does returning an empty Vec make sense when no error occured?

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

    pub fn render_user_context(&self) -> Result<Vec<EventVariantValue>, String> {
        let (raw_buffer, property_count) =
            self.render_generic(&[], EvtRenderContextUser, EvtRenderEventValues)?;
        let buffer = unsafe { EventVariantBuffer::from_raw_buffer(raw_buffer, property_count) };
        Ok(buffer.into_iter().collect())
    }

    pub fn render_xml(&self) -> Result<String, String> {
        let (raw_buffer, _) = self.render_generic(&[], 0, EvtRenderEventXml)?;

        Ok((raw_buffer.as_ptr() as *const u16).win_into())
    }

    fn render_message_int(
        &self,
        metadata: isize,
        max_buf_len: usize,
    ) -> Result<String, (WindowsError, u32)> {
        let mut message_buf: Vec<u16> = vec![0; max_buf_len];
        let mut buffer_used: u32 = 0;

        let result = unsafe {
            EvtFormatMessage(
                metadata,
                *self.handle,
                0,
                0,
                null(),
                EvtFormatMessageEvent,
                message_buf.len() as u32,
                message_buf.as_mut_ptr(),
                &mut buffer_used,
            )
        };

        if result != 0 {
            Ok(message_buf.as_ptr().win_into())
        } else {
            Err((WindowsError::from_win32(), buffer_used))
        }
    }

    pub fn render_message(&self) -> Result<String, String> {
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

        match buffer.get_property_value(1) {
            Some(EventVariantValue::String(str)) => return Ok(str),
            Some(EventVariantValue::Null) => {
                let provider_name = HSTRING::from(match buffer.get_property_value(0) {
                    Some(EventVariantValue::String(str)) => str,
                    _ => return Err("Unexpected result on provider name query".to_owned()),
                });

                let metadata_handle = unsafe {
                    EvtOpenPublisherMetadata(
                        0 as EVT_HANDLE,
                        provider_name.as_ptr(),
                        null(),
                        0, // LANG_NEUTRAL and SORT_DEFAULT
                        0,
                    )
                };

                let result = self.render_message_int(metadata_handle, 512);

                let (error, buffer_size) = match result {
                    Ok(str) => return Ok(str),
                    Err((err, bs)) => (err, bs),
                };

                if (error.code() != HRESULT::from_win32(ERROR_INSUFFICIENT_BUFFER)) {
                    return Err(format!(
                        "Error during initial message formatting: {:?}",
                        error.code()
                    ));
                }

                let result = self.render_message_int(metadata_handle, buffer_size as usize);

                match result {
                    Ok(str) => return Ok(str),
                    Err((error, _)) => {
                        return Err(format!(
                            "Error during message formatting: {:?}",
                            error.code()
                        ));
                    }
                }
            }
            _ => return Err("Unexpected result on message query".to_owned()),
        };
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

    pub fn get_property_value(&self, index: u32) -> Option<EventVariantValue> {
        if index >= self.property_count {
            return None;
        }

        let raw_variant = unsafe { self.index(index as isize) };

        Some((*raw_variant).into())
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
    type Item = EventVariantValue;

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
    type Item = EventVariantValue;

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
    FileTime(DateTime<Utc>),
    SysTime(NaiveDateTime),
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
    FileTimeArr(Vec<DateTime<Utc>>),
    SysTimeArr(Vec<NaiveDateTime>),
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
    UnknownType(i32),
    UnknownTypeArr(i32),
}

fn format_guid(guid: &GUID) -> String {
    format!(
        "{{{:08x}-{:04x}-{:04x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}}}",
        guid.data1,
        guid.data2,
        guid.data3,
        guid.data4[0],
        guid.data4[1],
        guid.data4[2],
        guid.data4[3],
        guid.data4[4],
        guid.data4[5],
        guid.data4[6],
        guid.data4[7]
    )
}

fn format_sid(value: &SID) -> String {
    use std::slice;
    let revision = unsafe { *(&value.Revision as *const u8) };
    let sub_authority_count = unsafe { *(&value.SubAuthorityCount as *const u8) };
    let identifier_authority = &value.IdentifierAuthority.Value;
    let sub_authorities = unsafe {
        slice::from_raw_parts(
            &value.SubAuthority as *const u32,
            sub_authority_count as usize,
        )
    };

    let identifier_authority = if identifier_authority[0..5] == [0, 0, 0, 0, 0] {
        identifier_authority[5].to_string()
    } else {
        format!(
            "{}",
            u64::from_be_bytes([
                0,
                0,
                identifier_authority[0],
                identifier_authority[1],
                identifier_authority[2],
                identifier_authority[3],
                identifier_authority[4],
                identifier_authority[5]
            ])
        )
    };

    let sub_authorities = sub_authorities
        .iter()
        .map(|s| s.to_string())
        .collect::<Vec<_>>()
        .join("-");
    format!(
        "S-{}-{}-{}",
        revision, identifier_authority, sub_authorities
    )
}

impl fmt::Debug for EventVariantValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            EventVariantValue::Null => write!(f, "Null"),
            EventVariantValue::Bool(value) => f.debug_tuple("Bool").field(value).finish(),
            EventVariantValue::SByte(value) => f.debug_tuple("SByte").field(value).finish(),
            EventVariantValue::Int16(value) => f.debug_tuple("Int16").field(value).finish(),
            EventVariantValue::Int32(value) => f.debug_tuple("Int32").field(value).finish(),
            EventVariantValue::Int64(value) => f.debug_tuple("Int64").field(value).finish(),
            EventVariantValue::Byte(value) => f.debug_tuple("Byte").field(value).finish(),
            EventVariantValue::UInt16(value) => f.debug_tuple("UInt16").field(value).finish(),
            EventVariantValue::UInt32(value) => f.debug_tuple("UInt32").field(value).finish(),
            EventVariantValue::UInt64(value) => f.debug_tuple("UInt64").field(value).finish(),
            EventVariantValue::Single(value) => f.debug_tuple("Single").field(value).finish(),
            EventVariantValue::Double(value) => f.debug_tuple("Double").field(value).finish(),
            EventVariantValue::FileTime(value) => f.debug_tuple("FileTime").field(value).finish(),
            EventVariantValue::SysTime(value) => f.debug_tuple("SysTime").field(value).finish(),
            EventVariantValue::Guid(value) => write!(f, "Guid({})", format_guid(value)),
            EventVariantValue::HexInt32(value) => write!(f, "HexInt32(0x{:08X})", value),
            EventVariantValue::HexInt64(value) => write!(f, "HexInt64(0x{:016X})", value),
            EventVariantValue::String(value) => f.debug_tuple("String").field(value).finish(),
            EventVariantValue::AnsiString(value) => {
                f.debug_tuple("AnsiString").field(value).finish()
            }
            EventVariantValue::Binary(value) => write!(
                f,
                "Binary({:?})",
                value
                    .iter()
                    .map(|b| format!("{:02X}", b))
                    .collect::<Vec<_>>()
                    .join(" ")
            ),
            EventVariantValue::Sid(value) => write!(f, "Sid({})", format_sid(value)),
            EventVariantValue::SizeT(value) => f.debug_tuple("SizeT").field(value).finish(),
            EventVariantValue::BoolArr(value) => f.debug_tuple("BoolArr").field(value).finish(),
            EventVariantValue::SByteArr(value) => f.debug_tuple("SByteArr").field(value).finish(),
            EventVariantValue::Int16Arr(value) => f.debug_tuple("Int16Arr").field(value).finish(),
            EventVariantValue::Int32Arr(value) => f.debug_tuple("Int32Arr").field(value).finish(),
            EventVariantValue::Int64Arr(value) => f.debug_tuple("Int64Arr").field(value).finish(),
            EventVariantValue::ByteArr(value) => f.debug_tuple("ByteArr").field(value).finish(),
            EventVariantValue::UInt16Arr(value) => f.debug_tuple("UInt16Arr").field(value).finish(),
            EventVariantValue::UInt32Arr(value) => f.debug_tuple("UInt32Arr").field(value).finish(),
            EventVariantValue::UInt64Arr(value) => f.debug_tuple("UInt64Arr").field(value).finish(),
            EventVariantValue::SingleArr(value) => f.debug_tuple("SingleArr").field(value).finish(),
            EventVariantValue::DoubleArr(value) => f.debug_tuple("DoubleArr").field(value).finish(),
            EventVariantValue::FileTimeArr(value) => {
                f.debug_tuple("FileTimeArr").field(value).finish()
            }
            EventVariantValue::SysTimeArr(value) => {
                f.debug_tuple("SysTimeArr").field(value).finish()
            }
            EventVariantValue::GuidArr(value) => {
                let formatted: Vec<String> = value.iter().map(|g| format_guid(g)).collect();
                write!(f, "GuidArr({:?})", formatted)
            }
            EventVariantValue::HexInt32Arr(value) => {
                write!(
                    f,
                    "HexInt32Arr({:?})",
                    value
                        .iter()
                        .map(|v| format!("0x{:08X}", v))
                        .collect::<Vec<_>>()
                )
            }
            EventVariantValue::HexInt64Arr(value) => {
                write!(
                    f,
                    "HexInt64Arr({:?})",
                    value
                        .iter()
                        .map(|v| format!("0x{:016X}", v))
                        .collect::<Vec<_>>()
                )
            }
            EventVariantValue::StringArr(value) => f.debug_tuple("StringArr").field(value).finish(),
            EventVariantValue::AnsiStringArr(value) => {
                f.debug_tuple("AnsiStringArr").field(value).finish()
            }
            EventVariantValue::SidArr(value) => {
                let formatted: Vec<String> = value.iter().map(|s| format_sid(s)).collect();
                write!(f, "SidArr({:?})", formatted)
            }
            EventVariantValue::SizeTArr(value) => f.debug_tuple("SizeTArr").field(value).finish(),
            EventVariantValue::EvtHandle(value) => f.debug_tuple("EvtHandle").field(value).finish(),
            EventVariantValue::Xml(value) => f.debug_tuple("Xml").field(value).finish(),
            EventVariantValue::XmlArr(value) => f.debug_tuple("XmlArr").field(value).finish(),
            EventVariantValue::UnknownType(value) => {
                f.debug_tuple("UnknownType").field(value).finish()
            }
            EventVariantValue::UnknownTypeArr(value) => {
                f.debug_tuple("UnknownTypeArr").field(value).finish()
            }
        }
    }
}

impl From<EVT_VARIANT> for EventVariantValue {
    fn from(value: EVT_VARIANT) -> Self {
        let is_array = (value.Type & EVT_VARIANT_TYPE_ARRAY) != 0;
        let value_type = (value.Type & EVT_VARIANT_TYPE_MASK) as i32;
        let count = value.Count as usize;

        if is_array {
            unsafe {
                #![allow(nonstandard_style)]
                match value_type {
                    EvtVarTypeString => Self::StringArr(
                        from_raw_parts(value.Anonymous.StringArr as *const PCWSTR_SYS, count)
                            .iter()
                            .map(|s| (*s).win_into())
                            .collect(),
                    ),
                    EvtVarTypeAnsiString => Self::AnsiStringArr(
                        from_raw_parts(value.Anonymous.AnsiStringArr as *const PCSTR_SYS, count)
                            .iter()
                            .map(|s| (*s).win_into())
                            .collect(),
                    ),
                    EvtVarTypeSByte => {
                        Self::SByteArr(from_raw_parts(value.Anonymous.SByteArr, count).to_vec())
                    }
                    EvtVarTypeByte => {
                        Self::ByteArr(from_raw_parts(value.Anonymous.ByteArr, count).to_vec())
                    }
                    EvtVarTypeInt16 => {
                        Self::Int16Arr(from_raw_parts(value.Anonymous.Int16Arr, count).to_vec())
                    }
                    EvtVarTypeUInt16 => {
                        Self::UInt16Arr(from_raw_parts(value.Anonymous.UInt16Arr, count).to_vec())
                    }
                    EvtVarTypeInt32 => {
                        Self::Int32Arr(from_raw_parts(value.Anonymous.Int32Arr, count).to_vec())
                    }
                    EvtVarTypeUInt32 => {
                        Self::UInt32Arr(from_raw_parts(value.Anonymous.UInt32Arr, count).to_vec())
                    }
                    EvtVarTypeInt64 => {
                        Self::Int64Arr(from_raw_parts(value.Anonymous.Int64Arr, count).to_vec())
                    }
                    EvtVarTypeUInt64 => {
                        Self::UInt64Arr(from_raw_parts(value.Anonymous.UInt64Arr, count).to_vec())
                    }
                    EvtVarTypeSingle => {
                        Self::SingleArr(from_raw_parts(value.Anonymous.SingleArr, count).to_vec())
                    }
                    EvtVarTypeDouble => {
                        Self::DoubleArr(from_raw_parts(value.Anonymous.DoubleArr, count).to_vec())
                    }
                    EvtVarTypeBoolean => Self::BoolArr(
                        from_raw_parts(value.Anonymous.BooleanArr, count)
                            .iter()
                            .map(|b| *b != 0)
                            .collect(),
                    ),
                    EvtVarTypeGuid => {
                        Self::GuidArr(from_raw_parts(value.Anonymous.GuidArr, count).to_vec())
                    }
                    EvtVarTypeSizeT => {
                        Self::SizeTArr(from_raw_parts(value.Anonymous.SizeTArr, count).to_vec())
                    }
                    EvtVarTypeFileTime => Self::FileTimeArr(
                        from_raw_parts(value.Anonymous.FileTimeArr, count)
                            .iter()
                            .map(|f| (*f).win_into())
                            .collect(),
                    ),
                    EvtVarTypeSysTime => Self::SysTimeArr(
                        from_raw_parts(value.Anonymous.SysTimeArr, count)
                            .iter()
                            .map(|s| (*s).win_into())
                            .collect(),
                    ),
                    EvtVarTypeSid => Self::SidArr(
                        from_raw_parts(value.Anonymous.SidArr as *const *const SID, count)
                            .iter()
                            .map(|s| **s)
                            .collect(),
                    ),
                    EvtVarTypeHexInt32 => {
                        Self::HexInt32Arr(from_raw_parts(value.Anonymous.UInt32Arr, count).to_vec())
                    }
                    EvtVarTypeHexInt64 => {
                        Self::HexInt64Arr(from_raw_parts(value.Anonymous.UInt64Arr, count).to_vec())
                    }
                    EvtVarTypeEvtXml => Self::XmlArr(
                        from_raw_parts(value.Anonymous.XmlValArr, count)
                            .iter()
                            .map(|s| (*s as PCWSTR_SYS).win_into())
                            .collect(),
                    ),
                    _ => Self::UnknownTypeArr(value_type),
                }
            }
        } else {
            unsafe {
                #![allow(nonstandard_style)]
                match value_type {
                    EvtVarTypeNull => Self::Null,
                    EvtVarTypeString => Self::String(value.Anonymous.StringVal.win_into()),
                    EvtVarTypeAnsiString => {
                        Self::AnsiString(value.Anonymous.AnsiStringVal.win_into())
                    }
                    EvtVarTypeSByte => Self::SByte(value.Anonymous.SByteVal),
                    EvtVarTypeByte => Self::Byte(value.Anonymous.ByteVal),
                    EvtVarTypeInt16 => Self::Int16(value.Anonymous.Int16Val),
                    EvtVarTypeUInt16 => Self::UInt16(value.Anonymous.UInt16Val),
                    EvtVarTypeInt32 => Self::Int32(value.Anonymous.Int32Val),
                    EvtVarTypeUInt32 => Self::UInt32(value.Anonymous.UInt32Val),
                    EvtVarTypeInt64 => Self::Int64(value.Anonymous.Int64Val),
                    EvtVarTypeUInt64 => Self::UInt64(value.Anonymous.UInt64Val),
                    EvtVarTypeSingle => Self::Single(value.Anonymous.SingleVal),
                    EvtVarTypeDouble => Self::Double(value.Anonymous.DoubleVal),
                    EvtVarTypeBoolean => Self::Bool(value.Anonymous.BooleanVal != 0),
                    EvtVarTypeBinary => {
                        Self::Binary(from_raw_parts(value.Anonymous.BinaryVal, count).to_vec())
                    }
                    EvtVarTypeGuid => Self::Guid(Box::new(*value.Anonymous.GuidVal)),
                    EvtVarTypeSizeT => Self::SizeT(value.Anonymous.SizeTVal),
                    EvtVarTypeFileTime => Self::FileTime(value.Anonymous.FileTimeVal.win_into()),
                    EvtVarTypeSysTime => Self::SysTime((*value.Anonymous.SysTimeVal).win_into()),
                    EvtVarTypeSid => Self::Sid(Box::new(*(value.Anonymous.SidVal as *const SID))),
                    EvtVarTypeHexInt32 => Self::HexInt32(value.Anonymous.UInt32Val),
                    EvtVarTypeHexInt64 => Self::HexInt64(value.Anonymous.UInt64Val),
                    EvtVarTypeEvtHandle => Self::EvtHandle(value.Anonymous.EvtHandleVal),
                    EvtVarTypeEvtXml => Self::Xml(value.Anonymous.XmlVal.win_into()),
                    _ => Self::UnknownType(value_type),
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
