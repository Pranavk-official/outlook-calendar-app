const { Client } = require("@microsoft/microsoft-graph-client");

class CalendarService {
  constructor(accessToken) {
    this.client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
  }

  async createEvent(event) {
    return await this.client.api("/me/events").post(event);
  }

  async addAttendee(eventId, attendee) {
    return await this.client
      .api(`/me/events/${eventId}/attendees`)
      .post(attendee);
  }

  async deleteEvent(eventId) {
    return await this.client.api(`/me/events/${eventId}`).delete();
  }
}

module.exports = CalendarService;
