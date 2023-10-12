import 'dart:convert';
import 'dart:io';

import 'package:arb_excel_dor/src/assets.dart';
import 'package:excel/excel.dart';
import 'package:path/path.dart';

import 'arb.dart';

const _kRowHeader = 0;
const _kRowValue = 1;
const _kColCategory = 0;
const _kColText = 1;
const _kColDescription = 2;
const _kColValue = 3;

typedef TranslationsData = Map<String, dynamic>;

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

Translation parseExcel({
  required String filename,
  String sheetname = 'Text',
  int headerRow = _kRowHeader,
  int valueRow = _kRowValue,
}) {
  final buf = File(filename).readAsBytesSync();
  final excel = Excel.decodeBytes(buf);
  final sheet = excel.sheets[sheetname];
  if (sheet == null) {
    return Translation();
  }

  final List<ARBItem> items = [];
  final columns = sheet.rows[headerRow];
  String? category;
  for (int i = valueRow; i < sheet.rows.length; i++) {
    final row = sheet.rows[i];
    final currentCategory = row[_kColCategory]?.value?.toString();
    category ??= currentCategory;
    if (category != currentCategory) {
      category = currentCategory;
      final item = ARBItem(
        category: category,
        text: null,
        description: null,
        translations: {},
      );
      items.add(item);
    }

    final item = ARBItem(
      category: null,
      text: row[_kColText]?.value?.toString() ?? '',
      description: row[_kColDescription]?.value?.toString(),
      translations: {},
    );

    for (int i = _kColValue; i < sheet.maxCols; i++) {
      final lang = columns[i]?.value?.toString() ?? i.toString();
      item.translations[lang] = row[i]?.value?.toString() ?? '';
    }

    items.add(item);
  }

  final languages = columns
      .where((e) => e != null && e.colIndex >= _kColValue)
      .map<String>((e) => e?.value?.toString() ?? '')
      .toList();
  return Translation(languages: languages, items: items);
}

/// Writes a Excel file, includes all translations.
Future<void> writeExcel(
  List<String> filenames,
  List<TranslationsData> translationsDataList, {
  String sheetname = 'Text',
}) async {
  final excel = Excel.createExcel();

  excel.updateCell(sheetname, CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0), 'category');
  excel.updateCell(sheetname, CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 0), 'text');
  excel.updateCell(sheetname, CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0), 'description');

  excel.setDefaultSheet(sheetname);

  final defaultSheet = excel.sheets[sheetname];

  final List<String> dataKeys = _getKeys(translationsDataList);

  for (var i = 0; i < dataKeys.length; i++) {
    final key = dataKeys[i];

    if (key.contains('@@')) {
      _updateTranslationRow(defaultSheet, translationsDataList, key, i + 1);
      _updateTranslationRow(defaultSheet, translationsDataList, key, 0, updateKey: false);
    } else if (key.contains('@')) {
      defaultSheet!.updateCell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + 1), key);
    } else {
      _updateTranslationRow(defaultSheet, translationsDataList, key, i + 1);
    }

    defaultSheet!.setColAutoFit(1);
    for (var j = 0; j < translationsDataList.length; j++) {
      defaultSheet.setColAutoFit(j + 3);
    }
  }
  final file = filenames.length == 1 ? File('${withoutExtension(filenames.first)}.xlsx') : File('intl.xlsx');
  excel.delete('Sheet1');
  await file.writeAsBytes(excel.save(fileName: file.path)!);
}

void _updateTranslationRow(
  defaultSheet,
  translationsDataList,
  String key,
  int index, {
  bool updateKey = true,
}) {
  if (updateKey) defaultSheet!.updateCell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: index), key);
  for (var j = 0; j < translationsDataList.length; j++) {
    var value = translationsDataList[j][key];
    if (value != null) {
      defaultSheet.updateCell(CellIndex.indexByColumnRow(columnIndex: j + 3, rowIndex: index), value);
    }
  }
}

List<String> _getKeys(translationsDataList) {
  final List<String> dataKeys = [];
  for (var translationsData in translationsDataList) {
    final keys = translationsData.keys.toList(growable: false);
    for (var key in keys) {
      if (!dataKeys.contains(key)) dataKeys.add(key);
    }
  }
  return dataKeys;
}
