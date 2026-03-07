from __future__ import annotations

from datetime import date, time

from pydantic import BaseModel, Field


class CreatePerformanceRequest(BaseModel):
    theatre_name: str
    theatre_city: str
    show_name: str
    performance_date: date
    performance_time: time
    supplier_reference: str | None = None


class CreatePerformanceResponse(BaseModel):
    performance_id: int
    theatre_id: int


class ImportTicketStockRequest(BaseModel):
    performance_id: int
    csv_content: str


class ImportBookingsRequest(BaseModel):
    performance_id: int
    csv_content: str


class RunAllocationRequest(BaseModel):
    performance_id: int


class ManualAllocationRequest(BaseModel):
    assigned_seats: list[str] = Field(default_factory=list)


class LoadSeatPlanRequest(BaseModel):
    theatre_id: int
    csv_content: str


class AllocationRowResponse(BaseModel):
    booking_id: int
    booking_reference: str
    customer_name: str
    quantity: int
    assigned_seats: list[str]
    section: str | None = None
    match_status: str
    match_notes: str
    manually_edited: bool
