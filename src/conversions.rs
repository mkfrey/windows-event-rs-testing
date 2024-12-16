
use windows_strings::{PCSTR, PCWSTR};
use windows_sys::Win32::Foundation::FILETIME;
use windows_sys::{
    core::{PCSTR as PCSTR_SYS, PCWSTR as PCWSTR_SYS},
    Win32::{Foundation::SYSTEMTIME, System::EventLog::EVT_HANDLE},
};

use chrono::{DateTime, NaiveDate, NaiveDateTime, Utc};
use std::time::{Duration, UNIX_EPOCH};

pub trait WindowsConversionFrom<T> {
    fn win_from(value: T) -> Self;
}

pub trait WindowsConversionTo<T> {
    fn win_into(self) -> T;
}

impl<T, U> WindowsConversionTo<U> for T
where
    U: WindowsConversionFrom<T>,
{
    fn win_into(self) -> U {
        U::win_from(self)
    }
}

impl WindowsConversionFrom<PCWSTR_SYS> for String {
    fn win_from(value: PCWSTR_SYS) -> Self {
        unsafe { PCWSTR::from_raw(value).to_string().unwrap() }
        //unsafe { PCWSTR::from_raw(value).to_hstring() }.to_string_lossy()
    }
}

impl WindowsConversionFrom<PCSTR_SYS> for String {
    fn win_from(value: PCSTR_SYS) -> Self {
        unsafe { PCSTR::from_raw(value).to_string().unwrap() }
    }
}

impl WindowsConversionFrom<u64> for DateTime<Utc> {
    /// Value is a windows timestamp containing number of elapsed 100 nsecs from Jan 1 1601
    fn win_from(value: u64) -> Self {
        let offset_to_unix: u64 = 116444736000000000;
        let unix_value = value - offset_to_unix;
        let nanos = unix_value.checked_mul(100).unwrap();
        DateTime::from_timestamp_nanos(nanos.try_into().unwrap())
    }
}

impl WindowsConversionFrom<u64> for FILETIME {
    fn win_from(value: u64) -> Self {
        Self {
            dwLowDateTime: (value & 0xFFFFFFFF) as u32,
            dwHighDateTime: (value >> 32) as u32,
        }
    }
}

impl WindowsConversionFrom<SYSTEMTIME> for NaiveDateTime {
    fn win_from(value: SYSTEMTIME) -> Self {
        NaiveDate::from_ymd_opt(value.wYear as i32, value.wMonth as u32, value.wDay as u32)
            .and_then(|date| {
                date.and_hms_milli_opt(
                    value.wHour as u32,
                    value.wMinute as u32,
                    value.wSecond as u32,
                    value.wMilliseconds as u32,
                )
            })
            .expect("Invalid SYSTEMTIME value")
    }
}

impl WindowsConversionFrom<FILETIME> for DateTime<Utc> {
    fn win_from(file_time: FILETIME) -> Self {
        let filetime_as_u64 =
            ((file_time.dwHighDateTime as u64) << 32) | (file_time.dwLowDateTime as u64);

        let epoch_difference = Duration::from_secs(11_644_473_600);
        let system_time =
            UNIX_EPOCH + Duration::from_nanos(filetime_as_u64 * 100) - epoch_difference;

        DateTime::<Utc>::from(system_time)
    }
}

/*
impl WindowsConversionFrom<SYSTEMTIME> for DateTime {
    fn win_from(value: SYSTEMTIME) -> Self {
        let date_time = DateTime::default();
        date_time.with_year(value.wYear)
        .and_then(|dt|  dt.with_month(value.wMonth))
        .and_then(|dt| dt.with_month(value.wMonth))
        .and_then(|dt| dt.with_day(value.wDay))
        .and_then(|dt| dt.with_hour(value.wHour))
        .and_then(|dt| dt.with_minute(value.wMinute))
        .and_then(|dt| dt.with_second(value.wSecond))
        .and_then(|dt| dt.mill(value.wMilliseconds))
 T   }
}*/
