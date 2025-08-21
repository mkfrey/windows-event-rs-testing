use std::ffi::c_void;
use std::ptr::null_mut;

use windows_result::Error as WindowsError;
use windows_result::HRESULT;
use windows_strings::HSTRING;
use windows_strings::PWSTR;
use windows_sys::Win32::Foundation::ERROR_NO_MORE_ITEMS;
use windows_sys::Win32::Foundation::S_OK;
use windows_sys::Win32::Foundation::WAIT_OBJECT_0;
use windows_sys::Win32::Foundation::{FALSE, TRUE};
use windows_sys::Win32::System::Com::*;
use windows_sys::Win32::System::EventLog::*;
use windows_sys::Win32::System::Threading::CreateEventW;
use windows_sys::Win32::System::Threading::ResetEvent;
use windows_sys::Win32::System::Threading::WaitForSingleObject;
use windows_sys::Win32::System::Threading::INFINITE;

mod conversions;
mod model;
mod tests;

use model::*;

static NULL_EVT_HANDLE: EVT_HANDLE = 0 as EVT_HANDLE;

pub extern "system" fn evt_subscribe_callback(
    action: EVT_SUBSCRIBE_NOTIFY_ACTION,
    _user_context: *const c_void,
    event_handle: EVT_HANDLE,
) -> u32 {
    match action {
        EvtSubscribeActionDeliver => (),
        _ => {
            println!("Received unexpected notify action {:?}", action);
            return 0; // Return value is ignored by caller
        }
    }

    let event: BorrowedWindowsEventHandle = BorrowedWindowsEventHandle::new(&event_handle);

    let system_context = event.render_system_context();

    match system_context {
        Ok(context) => {
            println!("Context Provider Name: {:?}", context.provider_name);
            println!("Context Event Id: {:?}", context.event_id);
            println!("Context Process Id: {:?}", context.process_id);
            match context.provider_guid {
                Some(guid) => {
                    println!("Context Provider GUID: {:?}", format_guid(&guid));
                    let mut winstr: *mut u16 = null_mut();
                    println!("Test");
                    let res = unsafe { StringFromCLSID(&guid, &mut winstr) };
                    if res != S_OK {
                        println!("Error converting GUID to string: {}", res);
                    } else {
                        let mut_str: PWSTR = PWSTR::from_raw(winstr);
                        let guid_str: String = unsafe { mut_str.to_string().unwrap() };
                        println!("Resolved Context Provider GUID: {:?}", guid_str);
                    }
                }

                None => println!("Context Provider GUID: None"),
            }
        }
        Err(error) => println!("Error getting system context: {}", error),
    };

    let xml = event.render_xml();

    match xml {
        Ok(xml) => {
            println!("{:?}", xml);
        }
        Err(error) => println!("Error getting XML: {}", error),
    };

    match event.render_user_context() {
        Ok(variables) => {
            println!("Event has {} properties.", variables.len());

            variables.iter().for_each(|f| println!("{:?}", f));
        }
        Err(err) => println!("Error: {}", err),
    }

    match event.render_message() {
        Ok(val) => println!("Description: {}", val),
        Err(err) => println!("Error rendering message: {}", err),
    }

    return 0;
}

fn main() {
    println!("Hello, world!");

    let agomillis = 3600000;
    let handle: EVT_HANDLE;
    let channel = HSTRING::from("Application");
    let query = HSTRING::from(format!(
        "*[System[TimeCreated[timediff(@SystemTime) <= {}]]]",
        agomillis
    ));

    let context: *mut c_void = null_mut();
    let flags: u32 = EvtSubscribeStartAtOldestRecord; // Swith to EvtSubscribeStartAfterBookmark later on, once bookmarking is implemented
    let new_event_event = unsafe {
        CreateEventW(
            null_mut(),
            TRUE,       // Manual reset
            TRUE,       // Initial state is non-signaled
            null_mut(), // No name
        )
    };

    if new_event_event == null_mut() {
        let last_error = WindowsError::from_win32();
        println!("Failed to create event: {:?}", last_error.message());
        return;
    }

    unsafe {
        handle = EvtSubscribe(
            NULL_EVT_HANDLE,
            new_event_event,
            channel.as_ptr(),
            query.as_ptr(),
            NULL_EVT_HANDLE,
            context,
            None,
            flags,
        );

        if handle == 0 {
            let last_error = WindowsError::from_win32();
            println!(
                "Received unexpected error while subscribing to events: {:?}",
                last_error.message()
            );
        }
    }

    loop {
        println!("Waiting for events...");
        let wait_result = unsafe {
            WaitForSingleObject(
                new_event_event,
                INFINITE, // INFINITE
            )
        };

        let mut events_returned: u32 = 0;
        let mut buffer: Vec<EVT_HANDLE> = Vec::with_capacity(10);

        if wait_result == WAIT_OBJECT_0 {
            println!("Event signaled, processing events...");

            // The event was signaled, meaning new events are available
            while unsafe {
                EvtNext(
                    handle,
                    buffer.capacity() as u32,
                    buffer.as_mut_ptr(),
                    10000000,
                    0,
                    &mut events_returned,
                )
            } == TRUE
            {
                unsafe { buffer.set_len(events_returned as usize) };

                if events_returned == 0 {
                    break;
                }

                println!("-----------------------------------");
                println!("Received {} events", events_returned);
                println!("-----------------------------------");

                for event_handle in buffer.iter() {
                    let event = OwnedWindowsEventHandle::new(*event_handle);
                    let system_context = event.render_system_context();

                    match system_context {
                        Ok(context) => {
                            println!("Context Provider Name: {:?}", context.provider_name);
                            println!("Context Event Id: {:?}", context.event_id);
                            println!("Context Process Id: {:?}", context.process_id);
                            match context.provider_guid {
                                Some(guid) => {
                                    println!("Context Provider GUID: {:?}", format_guid(&guid));
                                }
                                None => println!("Context Provider GUID: None"),
                            }
                        }
                        Err(error) => println!("Error getting system context: {}", error),
                    };

                    let xml = event.render_xml();

                    match xml {
                        Ok(xml) => {
                            println!("{:?}", xml);
                        }
                        Err(error) => println!("Error getting XML: {}", error),
                    };

                    match event.render_user_context() {
                        Ok(variables) => {
                            println!("Event has {} properties.", variables.len());

                            variables.iter().for_each(|f| println!("{:?}", f));
                        }
                        Err(err) => println!("Error: {}", err),
                    }

                    match event.render_message() {
                        Ok(val) => println!("Description: {}", val),
                        Err(err) => println!("Error rendering message: {}", err),
                    }
                }
            }

            let last_error = WindowsError::from_win32();
            if last_error.code() != HRESULT::from_win32(ERROR_NO_MORE_ITEMS) {
                println!(
                    "EvtNext failed: {:?} ({:?})",
                    last_error.message(),
                    last_error.code()
                );
                break;
            }

            // Reset the event to wait for new events again
            if unsafe { ResetEvent(new_event_event) } == FALSE {
                let last_error = WindowsError::from_win32();
                println!(
                    "ResetEvent failed: {:?} ({:?})",
                    last_error.message(),
                    last_error.code()
                );
            }
        } else {
            let last_error = WindowsError::from_win32();
            println!(
                "WaitForSingleObject failed: {:?} ({:?})",
                last_error.message(),
                last_error.code()
            );
        }
    }

    //thread::sleep(Duration::from_secs(60));
}
