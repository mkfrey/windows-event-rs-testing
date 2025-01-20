#[cfg(test)]
#[test]
fn test_time_conversion() {
    use crate::conversions::WindowsConversionTo;
    use windows_sys::Win32::Foundation::FILETIME;

    let time = FILETIME {
        dwHighDateTime: 0xAAAAAAAA as u32,
        dwLowDateTime: 0x55555555 as u32,
    };

    let int = 0xAAAAAAAA55555555 as u64;
    let time_from_int: FILETIME = int.win_into();

    assert_eq!(time.dwHighDateTime, time_from_int.dwHighDateTime);
    assert_eq!(time.dwLowDateTime, time_from_int.dwLowDateTime);
}

#[test]
fn test_systemtime_to_naive_datetime() {
    use crate::conversions::WindowsConversionFrom;
    use chrono::prelude::{Datelike, NaiveDateTime, Timelike};
    use windows_sys::Win32::Foundation::SYSTEMTIME;
    let system_time = SYSTEMTIME {
        wYear: 2024,
        wMonth: 12,
        wDay: 16,
        wHour: 12,
        wMinute: 30,
        wSecond: 45,
        wMilliseconds: 123,
        wDayOfWeek: 0,
    };
    let naive_datetime: NaiveDateTime = NaiveDateTime::win_from(system_time);
    assert_eq!(naive_datetime.year(), 2024);
    assert_eq!(naive_datetime.month(), 12);
    assert_eq!(naive_datetime.day(), 16);
    assert_eq!(naive_datetime.hour(), 12);
    assert_eq!(naive_datetime.minute(), 30);
    assert_eq!(naive_datetime.second(), 45);
    assert_eq!(naive_datetime.and_utc().timestamp_subsec_millis(), 123);
}

#[test]
fn test_filetime_to_datetime() {
    use crate::conversions::WindowsConversionFrom;
    use chrono::{DateTime, Datelike, Timelike, Utc};
    use windows_sys::Win32::Foundation::FILETIME;

    let file_time = FILETIME {
        dwLowDateTime: 0xD53E8000,
        dwHighDateTime: 0x01D96D5E,
    };
    let datetime: DateTime<Utc> = DateTime::win_from(file_time);
    assert_eq!(datetime.year(), 2023);
    assert_eq!(datetime.month(), 04);
    assert_eq!(datetime.day(), 12);
    assert_eq!(datetime.hour(), 16);
    assert_eq!(datetime.minute(), 50);
    assert_eq!(datetime.second(), 5);
}
