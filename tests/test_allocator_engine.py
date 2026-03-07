import unittest

from allocator.allocator_engine import run_allocation
from allocator.models import AvailableSeat, Booking, BookingPreference, TheatreSeat


class AllocatorEngineTests(unittest.TestCase):
    def test_keeps_booking_together_and_unallocates_if_insufficient_block(self) -> None:
        bookings = [
            Booking(id=1, performance_id=1, booking_reference="B100", customer_name="A", quantity=3),
            Booking(id=2, performance_id=1, booking_reference="B101", customer_name="B", quantity=2),
        ]

        preferences = {
            1: BookingPreference(booking_id=1),
            2: BookingPreference(booking_id=2),
        }

        available = [
            AvailableSeat(1, "Stalls:A:1", "Stalls", "A", 1, "A1"),
            AvailableSeat(1, "Stalls:A:2", "Stalls", "A", 2, "A2"),
            AvailableSeat(1, "Stalls:A:4", "Stalls", "A", 4, "A4"),
            AvailableSeat(1, "Stalls:A:5", "Stalls", "A", 5, "A5"),
        ]

        theatre_seats = [
            TheatreSeat(1, "Stalls", "A", 1, "A1", is_aisle=True, x_position=-1.5, y_position=0),
            TheatreSeat(1, "Stalls", "A", 2, "A2", is_aisle=False, x_position=-0.5, y_position=0),
            TheatreSeat(1, "Stalls", "A", 4, "A4", is_aisle=False, x_position=0.5, y_position=0),
            TheatreSeat(1, "Stalls", "A", 5, "A5", is_aisle=True, x_position=1.5, y_position=0),
        ]

        result = run_allocation(bookings, preferences, available, theatre_seats)
        by_booking = {a.booking_id: a for a in result.allocations}

        self.assertEqual(by_booking[1].assigned_seats, [])
        self.assertEqual(len(by_booking[2].assigned_seats), 2)


if __name__ == "__main__":
    unittest.main()
