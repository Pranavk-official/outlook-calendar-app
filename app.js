const express = require("express");
const passport = require("./auth.js");
const CalendarService = require("./calendar.service");
const app = express();
const logger = require("morgan");
const session = require("express-session");
const refresh = require('passport-oauth2-refresh');


// Add session middleware
app.use(
  session({
    secret: "your_session_secret", // Replace with a real secret
    resave: false,
    saveUninitialized: false,
  }),
);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(passport.initialize());
app.use(passport.session());

app.use(logger("dev"));

app.get("/login", passport.authenticate("microsoft"));
app.get(
  "/auth",
  passport.authenticate("microsoft", {
    failureRedirect: "/login",
  }),  function (req, res) {
    console.log(req.user)
    res.redirect('/success');
  }
);

function ensureAuthenticated(req, res, next) {
  if (req.isAuthenticated()) { return next(); }
  res.redirect('/login');
}
app.get("/success", ensureAuthenticated, function (req, res) {
  res.send(JSON.stringify(req.user))
})

app.get("/refresh", (res, req)=> {
  refresh.requestNewAccessToken(
    'facebook',
    req.user.refreshToken,
    function (err, accessToken, refreshToken) {
      req.user.accessToken = accessToken;
      req.user.refreshToken = refreshToken;
      res.send(accessToken, refreshToken)
    },
  );
})

app.get("/calendar", async (req, res) => {
  console.log(req.user)
  if (req.user && req.user.accessToken) {
    const calendarService = new CalendarService(req.user.accessToken);
    const attendee = {
      emailAddress: {
        address: "rajatxmathew@gmail.com",
        name: "New Attendee",
      },
      type: "required",
    };
    // Create an event
    const event = {
      subject: "Team Meeting",
      start: {
        dateTime: "2024-09-15T10:00:00",
        timeZone: "UTC",
      },
      end: {
        dateTime: "2024-09-15T11:00:00",
        timeZone: "UTC",
      },
      attendees: [
        {
          emailAddress: {
            address: "rajatmathew10@outlook.com",
            name: "Attendee Name",
          },
          type: "required",
        },
        // attendee
      ],
    };
    const response = await calendarService.createEvent(event);
    // Add an attendee to an event
    const eventId = response.id;

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
