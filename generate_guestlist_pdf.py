import csv
from collections import defaultdict
from reportlab.lib.pagesizes import A1, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch

def read_guests(csv_filename):
    guests = []
    with open(csv_filename, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            guests.append({
                'first_name': row.get('First Name', row.get('first_name', '')).strip(),
                'last_name': row.get('Last Name', row.get('last_name', '')).strip(),
                'table': row.get('Table', row.get('table', '')).strip(),
            })
    return guests

def group_by_alphabet(guests):
    groups = defaultdict(list)
    for guest in guests:
        letter = guest['last_name'][:1].upper() if guest['last_name'] else '#'
        groups[letter].append(guest)
    return dict(sorted(groups.items()))

def generate_guestlist_pdf(guests, filename):
    doc = SimpleDocTemplate(
        filename,
        pagesize=landscape(A1),
        leftMargin=0.5*inch,
        rightMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch,
        title="Wedding Guest List",
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name='GuestName',
        fontSize=28,
        leading=34,
        spaceAfter=8,
        alignment=0,
        fontName='Helvetica-Bold',
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
            name = f"{guest['first_name']} {guest['last_name']}"
            table = guest['table']
            content.append(Paragraph(f"{name} <font color='#888888'>{table}</font>", styles['GuestName']))
        content.append(Spacer(1, 32))

    doc.build(content)
    print(f"Generated {filename}")

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description="Generate a beautiful A1 PDF guestlist from seating.csv")
    parser.add_argument('--input', default='seating.csv', help='Input CSV file')
    parser.add_argument('--output', default='guestlist.pdf', help='Output PDF file')
    args = parser.parse_args()

    guests = read_guests(args.input)
    guests = sorted(guests, key=lambda x: (x['last_name'].upper(), x['first_name'].upper()))
    generate_guestlist_pdf(guests, args.output)