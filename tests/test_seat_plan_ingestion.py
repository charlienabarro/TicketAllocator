import unittest

from allocator.seat_plan_ingestion import SeatPlanIngestor, SeatPlanRow


class SeatPlanIngestionTests(unittest.TestCase):
    def test_infers_aisles_and_positions(self) -> None:
        ingestor = SeatPlanIngestor()
        result = ingestor.ingest_structured_rows(
            theatre_id=10,
            theatre_name="London Palladium",
            city="London",
            rows=[
                SeatPlanRow(section="Stalls", row="C", seat_number=1),
                SeatPlanRow(section="Stalls", row="C", seat_number=2),
                SeatPlanRow(section="Stalls", row="C", seat_number=4),
                SeatPlanRow(section="Stalls", row="C", seat_number=5),
            ],
        )

        self.assertEqual(len(result.theatre_seats), 4)
        by_num = {s.seat_number: s for s in result.theatre_seats}

        self.assertTrue(by_num[1].is_aisle)
        self.assertTrue(by_num[2].is_aisle)
        self.assertTrue(by_num[4].is_aisle)
        self.assertTrue(by_num[5].is_aisle)
        self.assertTrue(result.requires_manual_review)

    def test_unstructured_draft_requires_review(self) -> None:
        ingestor = SeatPlanIngestor()
        draft = ingestor.ingest_unstructured_text("Stalls C1 C2 C3 near aisle")
        self.assertTrue(draft.requires_manual_review)
        self.assertGreaterEqual(len(draft.tokens), 3)


if __name__ == "__main__":
    unittest.main()
