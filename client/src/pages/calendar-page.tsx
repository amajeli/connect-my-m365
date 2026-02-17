import { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { format, parseISO, isSameDay, startOfWeek, addDays } from "date-fns";
import type { OutlookCalendarEvent, AccountInfo } from "@shared/schema";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import { ScrollArea } from "@/components/ui/scroll-area";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Calendar,
  Clock,
  MapPin,
  Users,
  Video,
  ChevronLeft,
  ChevronRight,
  AlertCircle,
  Copy,
  Check,
  RefreshCw,
  Mail,
} from "lucide-react";
import { useToast } from "@/hooks/use-toast";

function getEventColor(index: number): string {
  const colors = [
    "bg-blue-500/10 border-blue-500/30 text-blue-700 dark:text-blue-300",
    "bg-purple-500/10 border-purple-500/30 text-purple-700 dark:text-purple-300",
    "bg-green-500/10 border-green-500/30 text-green-700 dark:text-green-300",
    "bg-orange-500/10 border-orange-500/30 text-orange-700 dark:text-orange-300",
    "bg-pink-500/10 border-pink-500/30 text-pink-700 dark:text-pink-300",
  ];
  return colors[index % colors.length];
}

export default function CalendarPage() {
  const [selectedDate, setSelectedDate] = useState(new Date());
  const [accountFilter, setAccountFilter] = useState("all");
  const [copied, setCopied] = useState(false);
  const { toast } = useToast();

  const { data: accounts } = useQuery<AccountInfo[]>({
    queryKey: ["/api/accounts"],
  });

  const {
    data: events,
    isLoading,
    isError,
    refetch,
    isFetching,
  } = useQuery<OutlookCalendarEvent[]>({
    queryKey: ["/api/calendar"],
    enabled: !!accounts && accounts.length > 0,
  });

  const weekStart = startOfWeek(selectedDate, { weekStartsOn: 0 });
  const weekDays = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));

  const filteredEvents = events?.filter((event) => {
    const matchesAccount = accountFilter === "all" || event.accountId === accountFilter;
    return matchesAccount;
  });

  const getEventsForDay = (day: Date) => {
    return filteredEvents?.filter((event) => {
      const eventStart = parseISO(event.start.dateTime);
      return isSameDay(eventStart, day);
    }) || [];
  };

  const todaysEvents = filteredEvents
    ?.filter((event) => {
      const eventStart = parseISO(event.start.dateTime);
      return isSameDay(eventStart, selectedDate);
    })
    .sort((a, b) => parseISO(a.start.dateTime).getTime() - parseISO(b.start.dateTime).getTime()) || [];

  const handleCopyForClaude = async () => {
    if (!todaysEvents.length) return;

    const eventsText = todaysEvents
      .map(
        (e) =>
          `Event: ${e.subject}\nTime: ${format(parseISO(e.start.dateTime), "h:mm a")} - ${format(parseISO(e.end.dateTime), "h:mm a")}\nLocation: ${e.location?.displayName || "N/A"}\n${e.isOnlineMeeting ? "Online Meeting" : ""}\nAttendees: ${e.attendees?.map((a) => a.emailAddress.name).join(", ") || "N/A"}\n---`
      )
      .join("\n\n");

    const text = `# Calendar for ${format(selectedDate, "EEEE, MMMM d, yyyy")}\n\n${eventsText}`;

    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
      toast({ title: "Copied to clipboard", description: "Paste this into Claude.ai" });
    } catch {
      toast({ title: "Failed to copy", variant: "destructive" });
    }
  };

  const goToPrevWeek = () => setSelectedDate(addDays(selectedDate, -7));
  const goToNextWeek = () => setSelectedDate(addDays(selectedDate, 7));
  const goToToday = () => setSelectedDate(new Date());

  if (!accounts || accounts.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-full p-8 text-center">
        <div className="w-16 h-16 rounded-full bg-primary/10 flex items-center justify-center mb-4">
          <Calendar className="w-8 h-8 text-primary" />
        </div>
        <h2 className="text-xl font-semibold mb-2">Connect your Outlook</h2>
        <p className="text-sm text-muted-foreground max-w-md mb-6">
          Connect a Microsoft account to view your calendar events.
        </p>
        <a href="/api/auth/microsoft" data-testid="button-connect-calendar">
          <Button className="gap-2">
            <Mail className="w-4 h-4" />
            Connect Microsoft Account
          </Button>
        </a>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-full">
      <div className="p-3 border-b flex items-center gap-2 flex-wrap">
        <div className="flex items-center gap-1">
          <Button size="icon" variant="ghost" onClick={goToPrevWeek} data-testid="button-prev-week">
            <ChevronLeft className="w-4 h-4" />
          </Button>
          <Button variant="outline" size="sm" onClick={goToToday} data-testid="button-today">
            Today
          </Button>
          <Button size="icon" variant="ghost" onClick={goToNextWeek} data-testid="button-next-week">
            <ChevronRight className="w-4 h-4" />
          </Button>
        </div>
        <h2 className="text-sm font-semibold">
          {format(selectedDate, "MMMM yyyy")}
        </h2>
        <div className="flex-1" />
        {accounts.length > 1 && (
          <Select value={accountFilter} onValueChange={setAccountFilter}>
            <SelectTrigger className="w-[180px]" data-testid="select-calendar-account-filter">
              <SelectValue placeholder="All accounts" />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="all">All accounts</SelectItem>
              {accounts.map((acc) => (
                <SelectItem key={acc.id} value={acc.id}>
                  {acc.email}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        )}
        <Button
          size="icon"
          variant="ghost"
          onClick={() => refetch()}
          disabled={isFetching}
          data-testid="button-refresh-calendar"
        >
          <RefreshCw className={`w-4 h-4 ${isFetching ? "animate-spin" : ""}`} />
        </Button>
        <Button
          size="sm"
          variant="outline"
          onClick={handleCopyForClaude}
          className="gap-1.5"
          data-testid="button-copy-calendar"
        >
          {copied ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
          {copied ? "Copied" : "Copy for Claude"}
        </Button>
      </div>

      <div className="grid grid-cols-7 border-b">
        {weekDays.map((day) => (
          <button
            key={day.toISOString()}
            className={`p-2 text-center hover-elevate transition-colors border-r last:border-r-0 ${
              isSameDay(day, selectedDate) ? "bg-primary/10" : ""
            } ${isSameDay(day, new Date()) ? "font-bold" : ""}`}
            onClick={() => setSelectedDate(day)}
            data-testid={`button-day-${format(day, "yyyy-MM-dd")}`}
          >
            <p className="text-[10px] text-muted-foreground uppercase">
              {format(day, "EEE")}
            </p>
            <p className={`text-lg ${isSameDay(day, new Date()) ? "text-primary" : ""}`}>
              {format(day, "d")}
            </p>
            {getEventsForDay(day).length > 0 && (
              <div className="flex justify-center gap-0.5 mt-0.5">
                {getEventsForDay(day).slice(0, 3).map((_, i) => (
                  <div
                    key={i}
                    className="w-1 h-1 rounded-full bg-primary"
                  />
                ))}
              </div>
            )}
          </button>
        ))}
      </div>

      <ScrollArea className="flex-1">
        <div className="p-4">
          <h3 className="text-sm font-semibold mb-3" data-testid="text-selected-date">
            {format(selectedDate, "EEEE, MMMM d")}
            {todaysEvents.length > 0 && (
              <span className="text-muted-foreground font-normal ml-2">
                {todaysEvents.length} event{todaysEvents.length !== 1 ? "s" : ""}
              </span>
            )}
          </h3>

          {isLoading ? (
            <div className="space-y-3">
              {Array.from({ length: 4 }).map((_, i) => (
                <Skeleton key={i} className="h-20 w-full rounded-md" />
              ))}
            </div>
          ) : isError ? (
            <div className="flex flex-col items-center py-8 gap-3">
              <AlertCircle className="w-10 h-10 text-destructive" />
              <p className="text-sm text-muted-foreground">Failed to load calendar</p>
              <Button variant="outline" size="sm" onClick={() => refetch()}>
                Retry
              </Button>
            </div>
          ) : todaysEvents.length > 0 ? (
            <div className="space-y-2">
              {todaysEvents.map((event, idx) => (
                <Card
                  key={`${event.accountId}-${event.id}`}
                  className={`p-3 border-l-4 ${getEventColor(idx)}`}
                  data-testid={`card-event-${event.id}`}
                >
                  <div className="flex items-start justify-between gap-2 flex-wrap">
                    <h4 className="text-sm font-medium">{event.subject}</h4>
                    {accounts.length > 1 && (
                      <Badge variant="secondary" className="text-[9px]">
                        {event.accountEmail}
                      </Badge>
                    )}
                  </div>
                  <div className="flex items-center gap-4 mt-2 text-xs text-muted-foreground flex-wrap">
                    <span className="flex items-center gap-1">
                      <Clock className="w-3 h-3" />
                      {event.isAllDay
                        ? "All day"
                        : `${format(parseISO(event.start.dateTime), "h:mm a")} - ${format(parseISO(event.end.dateTime), "h:mm a")}`}
                    </span>
                    {event.location?.displayName && (
                      <span className="flex items-center gap-1">
                        <MapPin className="w-3 h-3" />
                        {event.location.displayName}
                      </span>
                    )}
                    {event.isOnlineMeeting && (
                      <span className="flex items-center gap-1">
                        <Video className="w-3 h-3" />
                        Online
                      </span>
                    )}
                    {event.attendees && event.attendees.length > 0 && (
                      <span className="flex items-center gap-1">
                        <Users className="w-3 h-3" />
                        {event.attendees.length} attendee{event.attendees.length !== 1 ? "s" : ""}
                      </span>
                    )}
                  </div>
                  {event.bodyPreview && (
                    <p className="text-xs text-muted-foreground mt-2 line-clamp-2">
                      {event.bodyPreview}
                    </p>
                  )}
                </Card>
              ))}
            </div>
          ) : (
            <div className="flex flex-col items-center py-12 text-center">
              <Calendar className="w-10 h-10 text-muted-foreground mb-3" />
              <p className="text-sm font-medium">No events</p>
              <p className="text-xs text-muted-foreground mt-1">
                Nothing scheduled for this day
              </p>
            </div>
          )}
        </div>
      </ScrollArea>
    </div>
  );
}
