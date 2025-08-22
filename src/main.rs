mod conversions;
mod model;
mod tests;

use model::*;

fn main() {
    let agomillis = 3600000;
    let channel = "Application";
    let query = format!(
        "*[System[TimeCreated[timediff(@SystemTime) <= {}]]]",
        agomillis
    );

    let subscription = WindowsEventLogPollingSubscription::new(channel, Some(query.as_str()), None)
        .unwrap_or_else(|err| {
            eprintln!("Failed to create subscription: {}", err);
            std::process::exit(1);
        });

    let bookmark: WindowsEventLogBookmark = WindowsEventLogBookmark::new().unwrap_or_else(|err| {
        eprintln!("Failed to create bookmark: {}", err);
        std::process::exit(1);
    });

    subscription.read_events_blocking(
        |event| {
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

            bookmark.update(event).unwrap_or_else(|err| {
                eprintln!("Failed to update bookmark: {}", err);
            });
        },
        10,
        0,
    );
}
