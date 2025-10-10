import csv
from collections import defaultdict
from reportlab.lib.pagesizes import A1, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import BaseDocTemplate, Paragraph, Spacer, Frame, PageTemplate
from reportlab.lib.units import inch

def read_guests(csv_filename):
    guests = []
    with open(csv_filename, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # The CSV contains a single "Name" column and a "Table No." column.
            full_name = row.get('Name', row.get('name', '')).strip()
            # Split into first and last name: last name is the token after the last space.
            if full_name:
                if ' ' in full_name:
                    last_space = full_name.rfind(' ')
                    first = full_name[:last_space].strip()
                    last = full_name[last_space+1:].strip()
                else:
                    first = ''
                    last = full_name
            else:
                first = ''
                last = ''

            table = (row.get('Table No.', row.get('Table', row.get('table', '')))).strip()

            guests.append({
                'first_name': first,
                'last_name': last,
                'table': table,
            })
    return guests

def group_by_alphabet(guests):
    groups = defaultdict(list)
    for guest in guests:
        letter = guest['last_name'][:1].upper() if guest['last_name'] else '#'
        groups[letter].append(guest)
    return dict(sorted(groups.items()))

def generate_guestlist_pdf(guests, filename, columns=6):
    pagesize = landscape(A1)
    left_margin = right_margin = top_margin = bottom_margin = 0.5 * inch
    usable_width = pagesize[0] - left_margin - right_margin
    usable_height = pagesize[1] - top_margin - bottom_margin

    # Create frames for multi-column layout
    column_gap = 0.25 * inch
    column_width = (usable_width - (columns - 1) * column_gap) / columns
    frames = []
    for col in range(columns):
        x = left_margin + col * (column_width + column_gap)
        y = bottom_margin
        frames.append(Frame(x, y, column_width, usable_height, leftPadding=6, rightPadding=6, topPadding=6, bottomPadding=6))

    doc = BaseDocTemplate(
        filename,
        pagesize=pagesize,
        title="Wedding Guest List",
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
    )

    doc.addPageTemplates([PageTemplate(id='multi', frames=frames)])

    styles = getSampleStyleSheet()
    # Starting font size and leading for guest names
    base_font_size = 27.0
    base_leading = 35.0
    styles.add(ParagraphStyle(
        name='GuestName',
        fontSize=base_font_size,
        leading=base_leading,
        spaceAfter=2,
        alignment=0,
        fontName='Helvetica',
    ))
    styles.add(ParagraphStyle(
        name='Letter',
        fontSize=38,
        leading=56,
        spaceAfter=0,
        textColor=colors.HexColor('#3E3E3E'),
        fontName='Helvetica-Bold',
    ))

    groups = group_by_alphabet(guests)

    def make_content():
        """Create and return a fresh list of flowables for building.

        Platypus flowables are stateful and consumed during doc.build, so we must
        regenerate them each time before building.
        """
        _content = []
        for i, (letter, group) in enumerate(groups.items()):
            _content.append(Paragraph(letter, styles['Letter']))
            for guest in sorted(group, key=lambda x: (x['last_name'], x['first_name'])):
                first = guest['first_name']
                last = guest['last_name']
                table = guest['table']

                if first:
                    display_name = f"{first} <b>{last}</b>"
                else:
                    display_name = f"<b>{last}</b>"

                _content.append(Paragraph(f"{display_name} <font color='#888888'>{table}</font>", styles['GuestName']))
            # Append a spacer between groups, but not after the last one
            if i != len(groups) - 1:
                _content.append(Spacer(1, 32))
        return _content

    # Build once at base font size to see how many pages are produced.
    doc.build(make_content())
    page_count = doc.page if hasattr(doc, 'page') else None

    # If more than 1 page, iteratively reduce font size by 0.001 until it fits on one page
    if page_count is None:
        print(f"Generated {filename} (page count: unknown)")
        return

    # Avoid infinite loop: set a minimum font size we allow
    min_font_size = 6.0
    font_size = base_font_size
    leading = base_leading

    if page_count > 1:
        # We'll rebuild repeatedly; preserve original doc template, but recreate doc each time
        iteration = 0
        prev_page_count = page_count
        while page_count > 1 and font_size > min_font_size:
            iteration += 1
            font_size = round(font_size - 0.01, 3)
            # Scale leading proportionally to the original ratio
            leading = round(base_leading * (font_size / base_font_size), 3)

            # Update style
            styles['GuestName'].fontSize = font_size
            styles['GuestName'].leading = leading

            # Recreate document and frames to ensure layout uses updated styles
            doc = BaseDocTemplate(
                filename,
                pagesize=pagesize,
                title="Wedding Guest List",
                leftMargin=left_margin,
                rightMargin=right_margin,
                topMargin=top_margin,
                bottomMargin=bottom_margin,
            )
            doc.addPageTemplates([PageTemplate(id='multi', frames=frames)])

            # Rebuild with fresh flowables to get updated page count
            doc.build(make_content())
            page_count = doc.page if hasattr(doc, 'page') else None

            # Print periodic progress (every 50 iterations) and when page_count changes
            if iteration % 50 == 0 or page_count != prev_page_count:
                print(f"Iteration {iteration}: page_count={page_count}, fontSize={font_size}, leading={leading}")
                prev_page_count = page_count

        if page_count is None:
            print(f"Generated {filename} (page count: unknown)")
        else:
            print(f"Generated {filename} ({page_count} pages) with fontSize={font_size}")
    else:
        print(f"Generated {filename} ({page_count} pages)")

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description="Generate a beautiful A1 PDF guestlist from seating.csv")
    parser.add_argument('--input', default='seating.csv', help='Input CSV file')
    parser.add_argument('--output', default='guestlist.pdf', help='Output PDF file')
    parser.add_argument('--columns', default=7, type=int, help='Number of columns per page')
    args = parser.parse_args()

    guests = read_guests(args.input)
    guests = sorted(guests, key=lambda x: (x['last_name'].upper(), x['first_name'].upper()))
    generate_guestlist_pdf(guests, args.output, columns=args.columns)