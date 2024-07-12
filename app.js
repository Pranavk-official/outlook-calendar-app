const express = require("express");
const passport = require("./auth.js");
const CalendarService = require("./calendar.service");
const app = express();

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(passport.initialize());
app.use(passport.session());

app.get("/login", passport.authenticate("microsoft-graph"));
app.get(
  "/callback",
  passport.authenticate("microsoft-graph", {
    successRedirect: "/calendar",
    failureRedirect: "/login",
  }),
);

app.get("/calendar", async (req, res) => {
  if (req.user && req.user.accessToken) {
    const calendarService = new CalendarService(req.user.accessToken);

    // Create an event
    const event = {
      subject: "Team Meeting",
      start: {
        dateTime: "2023-07-15T10:00:00",
        timeZone: "UTC",
      },
      end: {
        dateTime: "2023-07-15T11:00:00",
        timeZone: "UTC",
      },
      attendees: [
        {
          emailAddress: {
            address: "attendee@example.com",
            name: "Attendee Name",
          },
          type: "required",
        },
      ],
    };
    await calendarService.createEvent(event);

    // Add an attendee to an event
    const eventId = "event-id-here";
    const attendee = {
      emailAddress: {
        address: "new-attendee@example.com",
        name: "New Attendee",
      },
      type: "required",
    };
    await calendarService.addAttendee(eventId, attendee);

    // Delete an event
    await calendarService.deleteEvent(eventId);

    res.send("Calendar operations completed");
  } else {
    res.redirect("/login");
  }
});

app.listen(3000, () => {
  console.log("Server is running on port 3000");
});
