from typing import Optional, List, Dict, Self
from datetime import datetime
from pydantic import BaseModel, Field, PrivateAttr
from orgparse import load
import re

class OrgNode(BaseModel):
    heading: str
    todo: Optional[str] = None
    tags: Optional[List[str]] = None
    properties: Optional[Dict[str, str]] = None
    body: Optional[str] = None
    timestamp: Optional[datetime] = None
    clocks: Optional[List[str]] = None
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
                    lines.append(f"CLOCK: {clock}")
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
            timestamp_pattern = r'^\s*<\d{4}-\d{2}-\d{2}\s+[A-Za-z]{3}\s+\d{2}:\d{2}>\s*$'

            filtered_lines = [line for line in body_lines
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
                timestamp = datetime.strptime(str(timestamp[0]), "<%Y-%m-%d %a %H:%M>")
            else:
                timestamp = None

            org_node = OrgNode(
                heading=node.heading,
                todo=node.todo,
                tags=node.tags,
                properties=node.properties,
                body=node.body,
                timestamp=timestamp,
                clocks=[str(clock) for clock in node.clock] if node.clock else None,
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
