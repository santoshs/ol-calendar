from typing import Optional, List, Dict, Self
from datetime import datetime, timedelta
from pydantic import BaseModel, Field, PrivateAttr
from orgparse import load
import re
import pytz

class OrgClock(BaseModel):
    start_time: datetime
    end_time: Optional[datetime] = None
    duration: Optional[timedelta] = None

    def __init__(self, start: str, end: Optional[str] = None):
        """
        Initialize OrgClock with start and end times as strings.
        Convert these to timezone-aware timestamps and calculate the duration.

        :param start: Start time as a string (e.g., "2024-12-10 04:30").
        :param end: End time as a string (e.g., "2024-12-10 06:30") or None.
        """
        start_time = self._convert_to_datetime(start)
        end_time = self._convert_to_datetime(end) if end else None
        duration = (end_time - start_time) if end_time else None

        super().__init__(start_time=start_time, end_time=end_time, duration=duration)

    @staticmethod
    def _convert_to_datetime(org_timestamp: str) -> datetime:
        """
        Convert an Org-mode timestamp (e.g., "2024-12-10 04:30") into a timezone-aware datetime.
        :param org_timestamp: Org-mode timestamp string.
        :return: Timezone-aware datetime object.
        """
        local_tz = pytz.timezone('Asia/Kolkata')
        naive_dt = datetime.strptime(org_timestamp, "%Y-%m-%d %H:%M:%S")
        return local_tz.localize(naive_dt)

    def to_org_string(self) -> str:
        """
        Convert the OrgClock instance to an Org-mode CLOCK string.
        :return: Org-mode formatted CLOCK string.
        """
        start_str = self.start_time.strftime("[%Y-%m-%d %a %H:%M]")
        end_str = self.end_time.strftime("[%Y-%m-%d %a %H:%M]") if self.end_time else ""
        duration_str = f" => {self._format_duration()}" if self.duration else ""
        return f"CLOCK: {start_str}--{end_str}{duration_str}"

    def _format_duration(self) -> str:
        """
        Format the duration as H:MM.
        :return: Duration string in H:MM format.
        """
        total_minutes = int(self.duration.total_seconds() // 60)
        hours, minutes = divmod(total_minutes, 60)
        return f"{hours}:{minutes:02d}"


class OrgNode(BaseModel):
    heading: str
    todo: Optional[str] = None
    tags: Optional[List[str]] = None
    properties: Optional[Dict[str, str]] = None
    body: Optional[str] = None
    timestamp: Optional[datetime] = None
    clocks: Optional[List[OrgClock]] = None
    level: int = 1

    def change_todo_state(self, new_state: Optional[str]) -> None:
        """
        Change the TODO state of the node.
        :param new_state: New TODO state (e.g., "TODO", "DONE", or None).
        """
        self.todo = new_state

    def update_timestamp(self, new_timestamp: datetime) -> None:
        """Update the timestamp of the node."""
        self.timestamp = new_timestamp

    def add_tag(self, tag: str) -> None:
        """Add a tag to the node if it does not already exist."""
        if self.tags is None:
            self.tags = []
        if tag not in self.tags:
            self.tags.append(tag)

    def remove_tag(self, tag: str) -> None:
        """Remove a tag from the node if it exists."""
        if self.tags and tag in self.tags:
            self.tags.remove(tag)


    def to_org_string(self) -> str:
        """Convert the node to an Org-mode formatted string."""
        lines = []

        # Add the TODO state, heading, and tags
        if self.heading != "":
            if self.todo:
                lines.append(f"{'*' * self.level} {self.todo} {self.heading}")
            else:
                lines.append(f"{'*' * self.level} {self.heading}")

            if self.tags:
                lines[-1] += f"  :{':'.join(self.tags)}:"

            if self.timestamp:
                lines.append(f"<{self.timestamp.strftime('%Y-%m-%d %H:%M')}>")

            if self.clocks:
                lines.append(":LOGBOOK:")
                for clock in self.clocks:
                    lines.append(clock.to_org_string())
                lines.append(":END:")

        # Add properties
        if self.properties:
            lines.append(":PROPERTIES:")
            for key, value in self.properties.items():
                lines.append(f":{key.upper()}: {value}")
            lines.append(":END:")

        # Add body
        if self.body:
            body_lines = self.body.splitlines()
            timestamp_pattern = r'^\s*<\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}>\s*$'

            filtered_lines = [f"{' ' * self.level + line.strip()}" for line in body_lines
                              if not re.match(timestamp_pattern, line)]

            if filtered_lines:
                lines.append('\n'.join(filtered_lines) + '\n')

        return "\n".join(lines)


class OrgFile(BaseModel):
    root: OrgNode = Field(...)
    children: List[OrgNode] = Field(...)

    @staticmethod
    def from_file(orgfile: str) -> Self:
        """Parse an Org file and create a list of OrgNode instances."""
        parsed_nodes = []
        root = load(orgfile)

        for node in root[1:]:  # Skip the root node
            timestamp = node.get_timestamps(active=True, point=True)
            if len(timestamp) > 0:
                naive_dt = datetime.strptime(str(timestamp[0]), "<%Y-%m-%d %a %H:%M>")
                india_tz = pytz.timezone('Asia/Kolkata')
                timestamp = india_tz.localize(naive_dt)
            else:
                timestamp = None

            # remove cancelled
            heading = node.heading
            todo = node.todo
            if heading.startswith("CANCELLED"):
                heading = heading[9:].strip()
                todo = "CANCELLED"

            org_node = OrgNode(
                heading=heading,
                todo=todo,
                tags=node.tags,
                properties=node.properties,
                body=node.body,
                timestamp=timestamp,
                clocks=[OrgClock(start=str(clock.start), end=str(clock.end)) for clock in node.clock] if node.clock else None,
                level = node.level
            )
            parsed_nodes.append(org_node)

        root_node = OrgNode(
            heading=root.heading,
            todo=None,
            tags=root.tags,
            properties=root.properties,
            body=root.body,
            level = 0
        )

        return OrgFile(root=root_node, children=parsed_nodes)

    def to_file(self, orgfile: str):
        with open(orgfile, "w") as f:
            f.write(self.root.to_org_string())
            for o in self.children:
                f.write(o.to_org_string())
