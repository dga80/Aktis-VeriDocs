import unittest
import os
import re
import tempfile
from PIL import Image

import config
import ocr_service
import veridocs_core

class TestVeriDocs(unittest.TestCase):
    def test_config_atomic_save(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            test_cfg_path = os.path.join(tmpdir, "test_config.json")
            orig_path = config.CONFIG_FILE
            try:
                config.CONFIG_FILE = test_cfg_path
                data = config.load_config()
                self.assertIn("duplex_mode", data)
                self.assertTrue(data["duplex_mode"])
                
                data["duplex_mode"] = False
                config.save_config(data)
                
                reloaded = config.load_config()
                self.assertFalse(reloaded["duplex_mode"])
            finally:
                config.CONFIG_FILE = orig_path

    def test_to_windows_long_path(self):
        # Normal local path
        p1 = r"C:\Scanner\output.pdf"
        lp1 = veridocs_core.to_windows_long_path(p1)
        self.assertTrue(lp1.startswith(r"\\?\C:\Scanner\output.pdf"))
        
        # UNC path
        p2 = r"\\server\share\output.pdf"
        lp2 = veridocs_core.to_windows_long_path(p2)
        self.assertTrue(lp2.startswith(r"\\?\UNC\server\share\output.pdf"))

    def test_blank_page_detection(self):
        # Create a pure white image
        white_img = Image.new("RGB", (800, 1000), (255, 255, 255))
        self.assertTrue(ocr_service.is_page_blank(white_img))

        # Create an image with significant dark content
        black_block_img = Image.new("RGB", (800, 1000), (255, 255, 255))
        for x in range(100, 700):
            for y in range(100, 700):
                black_block_img.putpixel((x, y), (0, 0, 0))
        self.assertFalse(ocr_service.is_page_blank(black_block_img))

    def test_of_regex_patterns(self):
        test_strings = [
            ("Orden de fabricacion AYC/MO/00560 en proceso", ["AYC/MO/00560"]),
            ("CESA/MO/00154 lote 2", ["CESA/MO/00154"]),
            ("Ficha LOESS|MO|00087 rev B", ["LOESS/MO/00087"]),
            ("Etiqueta MES/MO/05310 final", ["MES/MO/05310"]),
            ("Referencia MO/00524 urgente", ["MO/00524"]),
            ("Documento sin orden de fab", []),
        ]
        
        for text, expected in test_strings:
            detected = ocr_service._find_ofs_in_text(text)
            self.assertEqual(detected, expected, f"Failed on: {text}")

    def test_score_document_candidate(self):
        # Exact match
        score1 = ocr_service.score_document_candidate("PIE DE PAGINA DC-0021 REV A", "PIEDEPAGINADC-0021REVA", "DC-0021", "HOJA DE RUTA")
        self.assertEqual(score1, 100)

        # Spaced match (e.g. DC 0021)
        score2 = ocr_service.score_document_candidate("PIE DE PAGINA DC 0021 REV A", "PIEDEPAGINADC0021REVA", "DC-0021", "HOJA DE RUTA")
        self.assertEqual(score2, 90)

        # Compact match (e.g. DC0021)
        score3 = ocr_service.score_document_candidate("PIE DE PAGINA DC0021 REV A", "PIEDEPAGINADC0021REVA", "DC-0021", "HOJA DE RUTA")
        self.assertGreaterEqual(score3, 85)

        # Keyword match
        score4 = ocr_service.score_document_candidate("INFORME DE ENSAYOS FAT", "INFORMEDEENSAYOSFAT", "DC-0034", "FAT")
        self.assertGreaterEqual(score4, 65)

    def test_sanitize_filename(self):
        self.assertEqual(veridocs_core.sanitize_filename("test/file:name*?.pdf"), "test-file-name-.pdf")
        self.assertEqual(veridocs_core.format_of_for_filename("AYC/MO/00560"), "AYC_MO_00560")

    def test_sort_year_folders(self):
        folders = ["2023", "2025", "2024", "OLD", "HISTORICO"]
        sorted_folders = veridocs_core.sort_year_folders(folders)
        # 2025 and 2024 should come before 2023 and non-year folders
        self.assertEqual(sorted_folders[0], "2025")
        self.assertEqual(sorted_folders[1], "2024")
        self.assertEqual(sorted_folders[2], "2023")

    def test_monotonic_ids(self):
        session = veridocs_core.ScanningSession()
        session.log("Event 1")
        session.log("Event 2")
        status = session.get_status()
        self.assertEqual(len(status["events_log"]), 2)
        self.assertEqual(status["events_log"][0]["id"], 1)
        self.assertEqual(status["events_log"][1]["id"], 2)

    def test_processed_files_set_cleanup(self):
        session = veridocs_core.ScanningSession()
        with tempfile.NamedTemporaryFile(delete=False) as f:
            f.write(b"dummy")
            temp_path = f.name
        try:
            stat = os.stat(temp_path)
            sig = (temp_path, stat.st_mtime, stat.st_size)
            session.processed_files_set.add(sig)
            self.assertIn(sig, session.processed_files_set)
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

if __name__ == "__main__":
    unittest.main()
