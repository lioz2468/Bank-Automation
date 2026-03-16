import pdfplumber

print("=" * 80)
print("PDF FILE: 463560 גלובל.pdf")
print("=" * 80)

with pdfplumber.open('C:/Bank_Automation/463560 גלובל.pdf') as pdf:
    print(f"\nTotal pages: {len(pdf.pages)}")
    print(f"Metadata: {pdf.metadata}")
    print()

    for page_num, page in enumerate(pdf.pages, 1):
        print("=" * 80)
        print(f"PAGE {page_num}")
        print(f"  Width: {page.width}, Height: {page.height}")
        print()

        # Extract all text
        text = page.extract_text(x_tolerance=3, y_tolerance=3)
        print(f"  TEXT CONTENT:")
        if text:
            print(text)
        else:
            print("  (no text found)")
        print()

        # Extract words with positions
        words = page.extract_words(x_tolerance=3, y_tolerance=3, keep_blank_chars=True)
        print(f"  WORDS WITH POSITIONS ({len(words)} words):")
        for w in words:
            print(f"    x0={w['x0']:.2f} y0={w['top']:.2f} x1={w['x1']:.2f} y1={w['bottom']:.2f} text={repr(w['text'])}")
        print()

        # Extract tables
        tables = page.extract_tables()
        if tables:
            print(f"  TABLES ({len(tables)} found):")
            for t_idx, table in enumerate(tables, 1):
                print(f"    Table {t_idx}:")
                for row in table:
                    print(f"      {row}")
        else:
            print(f"  TABLES: none found")
        print()

        # Extract table with bounding boxes
        table_settings = {
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
        }
        tables_lines = page.extract_tables(table_settings)
        if tables_lines:
            print(f"  TABLES (lines strategy, {len(tables_lines)} found):")
            for t_idx, table in enumerate(tables_lines, 1):
                print(f"    Table {t_idx}:")
                for row in table:
                    print(f"      {row}")
        print()

        # Lines and rectangles
        lines = page.lines
        rects = page.rects
        print(f"  LINES: {len(lines)}")
        for l in lines:
            print(f"    {l}")
        print(f"  RECTS: {len(rects)}")
        for r in rects:
            print(f"    {r}")
        print()

print("=" * 80)
print("END OF PDF EXTRACTION")
print("=" * 80)
