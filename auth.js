const passport = require("passport");
const MicrosoftGraphStrategy = require("passport-microsoft").Strategy;
const refresh = require('passport-oauth2-refresh');

require("dotenv").config();
const strategy = new MicrosoftGraphStrategy(
  {
    clientID: process.env.MICROSOFT_APP_ID,
    clientSecret: process.env.MICROSOFT_APP_SECRET,
    callbackURL: process.env.MICROSOFT_REDIRECT_URI,
    scope: ["User.Read", "Calendars.ReadWrite"],
  },
  function (accessToken, refreshToken, profile, done) {
    // User profile is available in `profile` object
    return done(null, {accessToken, refreshToken ,...profile});
  },
);
passport.use(
  "microsoft",
  strategy
);
refresh.use(strategy);

passport.serializeUser(function (user, done) {
  done(null, user);
});

passport.deserializeUser(function (user, done) {
  done(null, user);
});

module.exports = passport;
