# Pilates Scheduling System – Google Apps Script Integration

This Google Apps Script connects Framer forms with Google Sheets and Google Calendar to automate client appointment management for a Pilates studio.

## Features

- Validates client subscriptions using the "Abonamente" sheet.
- Handles scheduling for 1, 2, or 3 sessions per week.
- Groups participants by experience level (Începător or Avansat).
- Prevents booking conflicts and mixed-level group sessions.
- Allows clients to reschedule sessions while maintaining constraints.
- Sends automated confirmation or rejection emails.
- Uses a single Web App endpoint with routing logic based on a hidden form field.

## Supported Request Types

Requests from Framer forms must include a hidden `type` field:
- `inscriere`: Registers a new client and adds data to the main sheet.
- `programare`: Books 1, 2, or 3 weekly recurring sessions.
- `reprogramare`: Reschedules a session to a new date/time.

## Sheet Structure

- `Abonamente`: Tracks client email addresses and remaining sessions.
- `Inscrieri`: New sign-up data.
- `Programari`, `Programari-2`, `Programari-3`: Scheduling requests based on subscription type.
- `Reprogramari`: Rescheduling requests.

## Technologies Used

- Google Apps Script
- Google Sheets API
- Google Calendar API
- Framer (frontend)

## Deployment Instructions

1. Deploy the script as a Web App (access: "Anyone" or "Anyone with the link").
2. Configure Framer forms to send POST requests to the Web App URL.
3. Include the appropriate `type` value to route to the correct handler function.
4. Ensure spreadsheet and calendar permissions are set correctly.

## Notes

- Booking is only allowed if the client has available sessions.
- Group classes must be consistent in experience level per time slot.
- Confirmation or rejection emails are automatically sent based on outcome.

---

This project streamlines scheduling operations for a Pilates studio while ensuring consistency and integrity across client bookings.
