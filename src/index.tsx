import child_process from "node:child_process";
import { promisify } from "node:util";
const exec = promisify(child_process.exec);
import { ActionPanel, List, Action, open } from "@raycast/api";
import { useCachedPromise } from "@raycast/utils";
import MicrosoftGraph from "@microsoft/microsoft-graph-types";
import assert from "node:assert";

type Event = Pick<
  MicrosoftGraph.Event,
  "subject" | "organizer" | "start" | "end" | "location" | "body" | "responseStatus"
> & { id: string };

const ONE_WEEK_MS = 7 * 24 * 60 * 60 * 1000;
const TIME_FORMAT_LANG = "en-US";
const TIME_FORMAT_OPTION = {
  month: "short",
  day: "2-digit",
  hour: "2-digit",
  minute: "2-digit",
  hour12: false,
} as const;

async function fetchEvents() {
  const now = new Date();
  const oneWeekLater = new Date(now.getTime() + ONE_WEEK_MS);

  const { stdout, stderr } = await exec(
    `mgc me calendar-view list \
      --headers 'Prefer=outlook.timezone="Asia/Tokyo"' \
      --select "subject,organizer,start,end,location,body,responseStatus" \
      --start-date-time ${now.toISOString()} \
      --end-date-time ${oneWeekLater.toISOString()} \
      --filter "isCancelled eq false" \
      --orderby "isAllDay,start/dateTime" \
      --top 50`,
  );

  if (stderr) throw new Error(stderr);

  const events: Event[] = JSON.parse(stdout).value;
  return events;
}

function filterEvents(events: Event[]): Event[] {
  const now = new Date();
  const endOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);

  const filteredEvents = events
    .filter((event) => {
      assert(event.start?.dateTime);
      assert(event.end?.dateTime);
      const start = new Date(event.start.dateTime);
      const end = new Date(event.end.dateTime);
      return end > now && start < endOfDay;
    })
    .filter((event) => event.responseStatus?.response !== "declined");

  return filteredEvents;
}

function extractLinks(event: Event): string[] {
  const content = event.location?.displayName + " " + event.body?.content;
  return [...new Set(content.match(/https?:\/\/[a-zA-Z0-9./?=_-]+/g))];
}

export default function Command() {
  const { isLoading, data } = useCachedPromise(fetchEvents);
  const events = filterEvents(data || []);

  return (
    <List isLoading={isLoading}>
      {events.map((event: Event) => {
        assert(event.start?.dateTime);
        assert(event.end?.dateTime);
        const displayStartTime = new Date(event.start.dateTime).toLocaleString(TIME_FORMAT_LANG, TIME_FORMAT_OPTION);
        const displayEndTime = new Date(event.end.dateTime).toLocaleString(TIME_FORMAT_LANG, TIME_FORMAT_OPTION);
        return (
          <List.Item
            key={event.id}
            icon="list-icon.png"
            title={event.subject || "No Title"}
            subtitle={`${displayStartTime} - ${displayEndTime}`}
            actions={
              <ActionPanel>
                <Action
                  title="Open Links"
                  onAction={() => {
                    extractLinks(event).forEach((link) => open(link));
                  }}
                />
              </ActionPanel>
            }
          />
        );
      })}
    </List>
  );
}
