import unittest

from allocator.booking_importer import parse_bookings_csv


class BookingImporterTests(unittest.TestCase):
    def test_normalizes_preference_text(self) -> None:
        csv_content = """booking_reference,customer_name,quantity,preferences,notes
B001,Jane Smith,2,aisle preferred,anniversary
B002,Mia Khan,3,"stalls only, must sit together",not too near the front
"""
        parsed = parse_bookings_csv(
            content=csv_content,
            performance_id=1,
            starting_booking_id=10,
            known_sections=["Stalls", "Royal Circle"],
        )

        self.assertEqual(len(parsed.bookings), 2)
        first_pref = parsed.preferences[0]
        self.assertTrue(first_pref.wants_aisle)

        second_pref = parsed.preferences[1]
        self.assertEqual(second_pref.section_preference, "Stalls")
        self.assertTrue(second_pref.section_preference_mandatory)
        self.assertTrue(second_pref.avoid_front)


if __name__ == "__main__":
    unittest.main()
