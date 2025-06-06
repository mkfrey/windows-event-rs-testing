use std::ffi::c_void;
use std::ptr::null_mut;
use std::thread;
use std::time::Duration;

use windows_result::Error as WindowsError;
use windows_strings::HSTRING;
use windows_sys::Win32::System::EventLog::*;

mod conversions;
mod model;
mod tests;

use conversions::*;
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

    let event = WindowsEvent::new(&event_handle);

    let system_context = event.render_system_context();

    match system_context {
        Ok(context) => {
            println!("{:?}", context.provider_name);
            println!("{:?}", context.event_id);
            println!("{:?}", context.process_id);
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

    unsafe {
        handle = EvtSubscribe(
            NULL_EVT_HANDLE,
            null_mut(),
            channel.as_ptr(),
            query.as_ptr(),
            NULL_EVT_HANDLE,
            context,
            Some(evt_subscribe_callback),
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

    thread::sleep(Duration::from_secs(60));
}
