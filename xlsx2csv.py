import click
import sys
from click.core import F
import openpyxl
import csv
import glob


def load_files(pattern, combine):
    files = glob.glob(pattern)
    files = [(f, openpyxl.load_workbook(f, data_only=True).active) for f in files if f.lower().endswith('.xlsx')]
    if not files:
        click.echo(f"No files found for pattern '{pattern}'")
        sys.exit(-1)
    return files


def write_separate(files, sep, encoding):
    for f in files:
        name, ws = f
        with open(name.replace('xlsx', 'csv'), 'w', encoding=encoding) as csv_file:
            writer = csv.writer(csv_file, delimiter=sep, quoting=csv.QUOTE_ALL)
            for row in ws.rows:
                writer.writerow([cell.value for cell in row])


def write_combined(files, output, sep, encoding):
    with open(output, 'w', encoding=encoding) as csv_file:
        writer = csv.writer(csv_file, delimiter=sep, quoting=csv.QUOTE_ALL)
        for i, f in enumerate(files):
            name, ws = f
            for j, row in enumerate(ws.rows):
                if i == 0 or j:
                    writer.writerow([cell.value for cell in row])


@click.command()
@click.option('-p', '--pattern', default='*.xlsx', help='Pattern to match input files, default *.xlsx')
@click.option('-d', '--delimiter', default=',', help='Default csv delimiter')
@click.option('-e', '--encoding', default='cp1252', help="The input file's encoding, default cp1252")
@click.option('-c', '--combined', default=False, is_flag=True, help='Combine all inputs into a single output file')
@click.option('-o', '--output', default='combined.csv', help='Output file for combined files')
def main(pattern, delimiter, encoding, combined, output):
    files = load_files(pattern, combined)
    if combined:
        write_combined(files, output, delimiter, encoding)
    else:
        write_separate(files, delimiter, encoding)

if __name__ == '__main__':
    main()
