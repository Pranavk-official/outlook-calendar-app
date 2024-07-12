const passport = require("passport");
const MicrosoftGraphStrategy = require("passport-microsoft").Strategy;
require("dotenv").config();

passport.use(
  new MicrosoftGraphStrategy(
    {
      clientID: process.env.MICROSOFT_APP_ID,
      clientSecret: process.env.MICROSOFT_APP_SECRET,
      callbackURL: process.env.MICROSOFT_REDIRECT_URI,
      scope: ["User.Read", "Calendars.ReadWrite"],
    },
    function (accessToken, refreshToken, profile, done) {
      // User profile is available in `profile` object
      return done(null, profile);
    },
  ),
);

passport.serializeUser(function (user, done) {
  done(null, user);
});

passport.deserializeUser(function (user, done) {
  done(null, user);
});

module.exports = passport;
