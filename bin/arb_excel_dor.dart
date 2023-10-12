import 'dart:io';

import 'package:args/args.dart';
import 'package:path/path.dart';
import 'package:arb_excel_dor/arb_excel_dor.dart';

const _kVersion = '0.0.1';

void main(List<String> args) async {
  final parse = ArgParser();
  parse.addFlag('arb', abbr: 'a', defaultsTo: false, help: 'Export to ARB files');
  parse.addFlag('excel', abbr: 'e', defaultsTo: false, help: 'Import ARB files to sheet');
  final flags = parse.parse(args);

  // Not enough args
  if (args.length < 2) {
    usage(parse);
    exit(1);
  }

  final filenames = flags.rest;

  if (flags['arb']) {
    final filename = filenames.first;
    stdout.writeln('Generate ARB from: $filename');
    final data = parseExcel(filename: filename);
    writeARB('${withoutExtension(filename)}.arb', data);
    exit(0);
  }

  if (flags['excel']) {
    stdout.writeln('Generate Excel from: $filenames');
    final List<Map<String, dynamic>> translations =
        await Future.wait(filenames.map((filename) async => await parseARB(filename)));
    stdout.writeln('ARB DATA PARSED');
    await writeExcel(filenames, translations);
    stdout.writeln('FINISHED');
    exit(0);
  }
}

void usage(ArgParser parse) {
  stdout.writeln('arb_sheet v$_kVersion\n');
  stdout.writeln('USAGE:');
  stdout.writeln(
    '  arb_sheet [OPTIONS] path/to/file/name\n',
  );
  stdout.writeln('OPTIONS');
  stdout.writeln(parse.usage);
}
