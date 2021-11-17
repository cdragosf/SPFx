import { graph } from "@pnp/graph/presets/all";
import { IEvent } from "../models/IEvent";
export class GraphCalendar {
    public async getTodayEvents(): Promise<IEvent[]> {
        let events: IEvent[] = [];
        const today = new Date();
        const endToday = new Date();
        endToday.setHours(23);
        endToday.setMinutes(59);


        const eventsResponse = await graph.me.calendarView(today.toISOString(), endToday.toISOString()).get();
        eventsResponse.map(event => {
            events.push(
                {
                    startTime: new Date(event.start.dateTime).toLocaleTimeString().slice(0,5),
                    endTime: new Date(event.end.dateTime).toLocaleTimeString().slice(0,5),
                    subject: event.subject,
                    url: event.webLink,
                    location: event.location.displayName
                }
            );
        });
        return events;
    }
}