# event-google-script

Script to synchronize birthdays and anniversaries from a spreadsheet.

## Usage
1. Create a Google Spreadsheet
1. Create Sheet named `Birthdays` with columns `Name` and `Birthday` (in `DD/MM` format), without any headers.
1. (Optional) Create Sheet named `Anniversaries` with columns `Name` and `Anniversary` (in `DD/MM` format), without any headers.
1. Go to `Extensions` -> `Apps Script` and paste the contents of `Code.gs` into the editor.
1. Create a `Google Calendar API` v3 service from `Services` section as `Calendar`.
1. Set a daily trigger to run the `syncEvents` function.

For more information on using Google Apps Script, refer to the [official documentation](https://developers.google.com/apps-script/guides/sheets).

## Notes

As of right now, Google Calendar API does not support creating `anniversary` events. So we reuse birthday events (with a different title) for anniversaries. 

See [events reference](https://developers.google.com/calendar/api/v3/reference/events#resource-representations) for more information.

> The Calendar API only supports creating events with the type "birthday". The type cannot be changed after the event is created.

## History

I first wrote this script around a decade back when I decided to delete my Facebook account but still keep track of birthdays on my calendar. It has run untouched since then, until now when I am refactoring it and putting it out here.

## Changelog
1. Sometime back: Initial version
1. 2025-11-01: Refactored code, use special birthday events instead of all-day events.