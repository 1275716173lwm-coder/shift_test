from __future__ import annotations

from abc import ABC, abstractmethod

from scheduler_app.core.models import Assignment, Employee


class HardRule(ABC):
    name: str

    @abstractmethod
    def validate(self, assignments: list[Assignment], candidate: Assignment, employees_by_id: dict[int, Employee]) -> tuple[bool, str]:
        ...


class SoftRule(ABC):
    name: str

    @abstractmethod
    def score(self, assignments: list[Assignment], candidate: Assignment, employees_by_id: dict[int, Employee]) -> int:
        ...
