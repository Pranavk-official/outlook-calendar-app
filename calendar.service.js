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
    const res = await this.client.api("/me/events/"+eventId).get();

    return await this.client
      .api(`/me/events/${eventId}`)
      .patch({
        attendees: [attendee, ...res.attendees]
      });
  }

  async deleteEvent(eventId) {
    return await this.client.api(`/me/events/${eventId}`).delete();
  }
}

module.exports = CalendarService;
