import 'dart:convert';
import 'dart:io';

import 'package:arb_excel/arb_excel.dart';
import 'package:path/path.dart';

/// To match all args from a text.
final _kRegArgs = RegExp(r'{(\w+)}');

/// Parses .arb files to [Translation].
/// The [filename] is the main language.

Future<Map<String, dynamic>> parseARB(String filename) async {
  try {
    final buf = await File(filename).readAsString();
    final map = jsonDecode(buf);
    return map;
  } catch (error) {
    print(error);
    throw Error();
  }
}

/// Writes [Translation] to .arb files.
void writeARB(String filename, Translation data) {
  for (var i = 0; i < data.languages.length; i++) {
    final lang = data.languages[i];
    final isDefault = i == 0;
    final f = File('${withoutExtension(filename)}_$lang.arb');

    List buf = [];
    for (final item in data.items) {
      final data = item.toJSON(lang, isDefault);
      if (data != null) {
        buf.add(item.toJSON(lang, isDefault));
      }
    }

    buf = ['{', buf.join(',\n'), '}\n'];
    f.writeAsStringSync(buf.join('\n'));
  }
}

/// Describes an ARB record.
class ARBItem {
  static List<String> getArgs(String text) {
    final List<String> args = [];
    final matches = _kRegArgs.allMatches(text);
    for (final m in matches) {
      final arg = m.group(1);
      if (arg != null) {
        args.add(arg);
      }
    }

    return args;
  }

  ARBItem({
    this.category,
    this.text,
    this.description,
    this.translations = const {},
  });

  final String? category;
  final String? text;
  final String? description;
  final Map<String, String> translations;

  /// Serialize in JSON.
  String? toJSON(String lang, [bool isDefault = false]) {
    final value = translations[lang];
    if (value == null || value.isEmpty) return null;

    final List<String> buf = [];
    if (category != null) {
      buf.add('  "@$category": {}');
    }
    if (text != null) {
      final data = {text: value};
      final jsonString = jsonEncode(data);
      final jsonSubstring = jsonString.substring(1, jsonString.length - 1);
      buf.add(jsonSubstring);
    }

    return buf.join('\n');
  }
}

/// Describes all arb records.
class Translation {
  Translation({this.languages = const [], this.items = const []});

  final List<String> languages;
  final List<ARBItem> items;
}
