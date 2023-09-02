import dataclasses
import datetime
from typing import Any

from win32com.client import Dispatch


@dataclasses.dataclass
class PIData:
    timestamp: str
    value: Any


@dataclasses.dataclass
class PITagValues:
    tag: str
    description: str
    eng_unit: str
    values: list[PIData]


class PI:
    def __init__(self, server: str) -> None:
        try:
            pisdk = Dispatch("PISDK.PISDK")
            self.server = pisdk.Servers(server)
            self.server_name = server
        except Exception as e:
            raise ConnectionError(
                "Server {} not found or could not connect! {}".format(server, e)
            )

    def read(
        self,
        tag: str,
        start: str | datetime.datetime | None = None,
        end: str | datetime.datetime | None = None,
        interval: int = 60,
    ) -> PITagValues:
        try:
            sample_point = self.server.PIPoints[tag]
        except Exception as e:
            raise ValueError(
                "TAG {} does not exist on {}! {}".format(tag, self.server_name, e)
            )

        if start is None and end is not None:
            raise ValueError("End date and time was provided but not start!")

        eng_unit = sample_point.PointAttributes.Item("EngUnits").Value
        description = sample_point.PointAttributes.Item("Descriptor").Value

        if start is None and end is None:
            value = sample_point.data.Snapshot.Value
            timestamp = sample_point.data.Snapshot.TimeStamp.LocalDate.strftime(
                "%Y-%m-%d %H:%M:%S"
            )
            return PITagValues(tag, description, eng_unit, [PIData(timestamp, value)])

        time_start = Dispatch("PITimeServer.PITimeFormat")
        time_end = Dispatch("PITimeServer.PITimeFormat")

        if isinstance(start, str):
            start = datetime.datetime.strptime(start, "%Y-%m-%d %H:%M:%S")
        elif not isinstance(start, datetime.datetime):
            raise ValueError("Start date and time must be str ou datetime object!")

        if end is not None:
            if isinstance(end, str):
                end = datetime.datetime.strptime(end, "%Y-%m-%d %H:%M:%S")
            elif not isinstance(end, datetime.datetime):
                raise ValueError("End date and time must be str ou datetime object!")
        else:
            end = start + datetime.timedelta(seconds=1)

        time_start.InputString = start.strftime("%Y-%m-%d %H:%M:%S")
        time_end.InputString = end.strftime("%Y-%m-%d %H:%M:%S")
        sample_values = sample_point.Data.Summaries2(
            time_start, time_end, interval, 5, 0, None
        )
        values = [
            PIData(x.TimeStamp.LocalDate.strftime("%Y-%m-%d %H:%M:%S"), x.Value)
            for x in sample_values("Average").Value
        ]

        return PITagValues(tag, description, eng_unit, values)
