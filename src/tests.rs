use crate::conversions::*;
use windows_sys::Win32::Foundation::FILETIME;


#[cfg(test)]

#[test]
fn test_time_conversion() {
    let time = FILETIME {
        dwHighDateTime: 0xAAAAAAAA as u32,
        dwLowDateTime:  0x55555555 as u32
    };

    let int = 0xAAAAAAAA55555555 as u64;
    let time_from_int: FILETIME = int.win_into();

    assert_eq!(time.dwHighDateTime, time_from_int.dwHighDateTime);
    assert_eq!(time.dwLowDateTime, time_from_int.dwLowDateTime);
}



#[test]
fn test_string_conversion() {

}