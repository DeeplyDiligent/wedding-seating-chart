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

def generate_guestlist_pdf(guests, filename, columns=4):
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
    styles.add(ParagraphStyle(
        name='GuestName',
        fontSize=27,
        leading=35,
        spaceAfter=2,
        alignment=0,
        fontName='Helvetica',
    ))
    styles.add(ParagraphStyle(
        name='Letter',
        fontSize=48,
        leading=56,
        spaceAfter=18,
        textColor=colors.HexColor('#3E3E3E'),
        fontName='Helvetica-Bold',
    ))

    content = []
    groups = group_by_alphabet(guests)
    for letter, group in groups.items():
        content.append(Paragraph(letter, styles['Letter']))
        for guest in sorted(group, key=lambda x: (x['last_name'], x['first_name'])):
            # Bold only the last name (the token after the last space in the original full name)
            first = guest['first_name']
            last = guest['last_name']
            table = guest['table']

            if first:
                # Use ReportLab's simple XML-like tags: <b> for bold and <font> for color
                display_name = f"{first} <b>{last}</b>"
            else:
                display_name = f"<b>{last}</b>"

            content.append(Paragraph(f"{display_name} <font color='#888888'>{table}</font>", styles['GuestName']))
        content.append(Spacer(1, 32))

    doc.build(content)
    print(f"Generated {filename}")

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description="Generate a beautiful A1 PDF guestlist from seating.csv")
    parser.add_argument('--input', default='seating.csv', help='Input CSV file')
    parser.add_argument('--output', default='guestlist.pdf', help='Output PDF file')
    parser.add_argument('--columns', default=4, type=int, help='Number of columns per page')
    args = parser.parse_args()

    guests = read_guests(args.input)
    guests = sorted(guests, key=lambda x: (x['last_name'].upper(), x['first_name'].upper()))
    generate_guestlist_pdf(guests, args.output, columns=args.columns)