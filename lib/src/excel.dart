import 'dart:convert';
import 'dart:io';

import 'package:arb_excel/src/assets.dart';
import 'package:excel/excel.dart';
import 'package:path/path.dart';

import 'arb.dart';

const _kRowHeader = 0;
const _kRowValue = 1;
const _kColCategory = 0;
const _kColText = 1;
const _kColDescription = 2;
const _kColValue = 3;

/// Create a new Excel template file.
///
/// Embedded data will be packed via `template.dart`.
void newTemplate(String filename) {
  final buf = base64Decode(kTemplate);
  File(filename).writeAsBytesSync(buf);
}

/// Reads Excel sheet.
///
/// Uses `arb_sheet -n path/to/file` to create a translation file
/// from the template.

// void main() async {
//   final translations = parseExcel(filename: 'example/example.xlsx');
//   print(translations.items.first);
//   writeARB('${withoutExtension('example/example.xlsx')}.arb', translations);
//   print(translations.items.first);
// }

Translation parseExcel({
  required String filename,
  String sheetname = 'Text',
  int headerRow = _kRowHeader,
  int valueRow = _kRowValue,
}) {
  final buf = File(filename).readAsBytesSync();
  final excel = Excel.decodeBytes(buf);
  Excel.createExcel();
  final sheet = excel.sheets[sheetname];
  if (sheet == null) {
    return Translation();
  }

  final List<ARBItem> items = [];
  final columns = sheet.rows[headerRow];
  for (int i = valueRow; i < sheet.rows.length; i++) {
    final row = sheet.rows[i];
    final item = ARBItem(
      category: row[_kColCategory]?.value?.toString(),
      text: row[_kColText]?.value?.toString() ?? '',
      description: row[_kColDescription]?.value?.toString(),
      translations: {},
    );

    try {
      for (int i = _kColValue; i < sheet.maxCols; i = i + 3) {
        final lang = columns[i]?.value?.toString() ?? i.toString();
        String translation = row[i]?.value?.toString() ?? '';
        if (translation.contains('[plural]')) {
          final one = row[i + 1]?.value?.toString();
          final few = row[i + 2]?.value?.toString();
          final other = row[i + 3]?.value?.toString();
          String plurals = '';
          if (one != null) {
            plurals += ' one{$one}';
          }
          if (few != null) {
            plurals += ' few{$few}';
          }
          if (other != null) {
            plurals += ' other{$other}';
          }
          translation = translation.replaceAll('[plural]', '{count,plural,$plurals}');
        }
        item.translations[lang] = translation;
      }
    } catch (error) {
      print(error);
    }

    items.add(item);
  }

  final languages = columns
      .where((e) => e != null && e.colIndex >= _kColValue && (e.colIndex + 1) % 4 == 0)
      .map<String>((e) => e?.value?.toString() ?? '')
      .toList();
  return Translation(languages: languages, items: items);
}

/// Writes a Excel file, includes all translations.
void writeExcel(
  String filename,
  Map<String, dynamic> data, {
  String sheetname = 'Text',
}) {
  final excel = Excel.createExcel();

  excel.updateCell(sheetname, CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0), 'category');
  excel.updateCell(sheetname, CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 0), 'text');
  excel.updateCell(sheetname, CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0), 'description');
  excel.updateCell(sheetname, CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: 0), 'en');

  excel.setDefaultSheet(sheetname);

  final defaultSheet = excel.sheets[sheetname];

  for (var i = 0; i < data.keys.length; i++) {
    final keys = data.keys.toList(growable: false);
    final key = keys[i];
    final value = data[key];

    if (key.contains('@')) {
      defaultSheet!.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), key);
    } else {
      defaultSheet!.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: i + 1), key);
      defaultSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: 3, rowIndex: i + 1), value);
    }

    defaultSheet.setColAutoFit(1);
    defaultSheet.setColAutoFit(3);
  }


  final file = File('example/${withoutExtension(filename)}.xlsx');
  excel.delete('Sheet1');
  file.writeAsBytes(excel.save(fileName: filename)!);
}
