// debug dart --pause-isolates-on-start --observe bin/cli.dart -f 'C:\Users\majona\Documents\Copy of 10 - Pointage Maintenance Octobre 2019.xlsx'
import 'dart:typed_data';
import 'package:args/args.dart';
import 'dart:io';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';
// as of march 31 excel does not phrase formulas so spreadsheet decoder is used instead
import 'package:intl/intl.dart';
import 'dart:convert';

final labors = {
  'HERISSE Olivier': {'laborcode': 'FRCOOHER', 'employee_number': 13},
  'ASTREINTES HERISSE': {'laborcode': 'FRCOOHER', 'employee_number': 13},
  'DUBOIS Michaël': {'laborcode': 'FRCOMDUB', 'employee_number': 133},
  'ASTREINTES DUBOIS M': {'laborcode': 'FRCOMDUB', 'employee_number': 133},
  'BAYRAND Pascal': {'laborcode': 'FRCOPBAY', 'employee_number': 50},
  'ASTREINTES BAYRAND P': {'laborcode': 'FRCOPBAY', 'employee_number': 50},
  'NATIVELLE Lucien': {'laborcode': 'FRCOLNAT', 'employee_number': 94},
  'ASTREINTES NATIVELLE L': {'laborcode': 'FRCOLNAT', 'employee_number': 94},
  'LEMOINE Angélis': {'laborcode': 'FRCOALEM', 'employee_number': 95},
  'SAVY Michaël': {'laborcode': 'FRCOMSAV', 'employee_number': 96},
};

void main(List<String> arguments) {
  exitCode = 0; //assume success
  final parser = ArgParser()
    ..addOption("file", help: "Path to file to convert", abbr: 'f');
  ArgResults argResults = parser.parse(arguments);
  var file = argResults["file"];
  if (file is! String) {
    stdout.writeln('Please specify filepath in arguments and try again');
    stdout.writeln('program.exe --file [filepath]');
    exit(2);
  }
  Uint8List bytes;
  try {
    bytes = File(file).readAsBytesSync();
  } catch (_) {
    stdout.writeln('Error reading file. Please try again');
    stdout.writeln('program.exe -f [filepath]');
    exit(2);
  }
  var decoder = SpreadsheetDecoder.decodeBytes(bytes, update: true);
  String newSheet =
      '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';
  // spreadsheet_decode doesnt actually have a function for new sheet
  // this is copied from excel
  var excel = SpreadsheetDecoder.decodeBytes(Base64Decoder().convert(newSheet),
      update: true);
  var sheet = excel.tables.keys.first;
  for (var i = 0; i <= 14; i++) {
    // rows and columns need to be predefined before usage
    excel.insertColumn(sheet, i);
  }
  excel.insertRow(sheet, 0); //blank row for header
  var headers = [
    'Empl#',
    'Employee Name',
    'Day',
    'Time In',
    'Time Out',
    'Hours',
    'LL1',
    'LL2',
    'LL3',
    'LL4',
    'LL5',
    'LL6',
    'LL7',
    'Reg',
    'O/T'
  ];
  var writeCol = 0;
  for (var cell in headers) {
    excel.updateCell(sheet, writeCol, 0, cell);
    writeCol++;
  }
  dynamic totalHrs;
  dynamic normalHrs;
  dynamic startTime;
  var writeRow = 1;
  writeCol = 0;
  for (var table in decoder.tables.keys) {
    if (!(labors.keys.contains(table))) {
      // only parse the sheet if the name is in the list
      print('skipping sheet "$table"');
      continue;
    }
    var employee = labors[table];
    for (var row in decoder.tables[table]!.rows) {
      if (row[2] is int) {
        if (table.contains('ASTREINTES')) {
          totalHrs = row[8];
          normalHrs = 0;
          startTime = row[4] ?? row[6] ?? 'Blank';
        } else {
          totalHrs = row[7];
          normalHrs = row[8];
          startTime = row[3] ?? row[5] ?? 'Blank';
        }
        if (isTimeStamp(startTime)) {
          var weekday = weekdays[row[1]];
          var startTimeStamp =
              addTime(toDartDateTime(row[2].toDouble()), startTime);
          var endTimeStamp = startTimeStamp.add(Duration(
            hours: totalHrs.toInt(),
            minutes: (totalHrs % 1 * 60).round(),
          ));
          var ot = totalHrs - normalHrs;
          var write = [
            employee?['employee_number'], // employee number
            employee?['laborcode'],
            weekday,
            DateFormat('yyyy-MM-dd HH:mm').format(startTimeStamp),
            DateFormat('yyyy-MM-dd HH:mm').format(endTimeStamp),
            (totalHrs).toStringAsFixed(2), // hours worked
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            (totalHrs < normalHrs ? totalHrs : normalHrs).toStringAsFixed(2),
            (ot > 0 ? ot : 0).toStringAsFixed(2)
          ];
          excel.insertRow(sheet, writeRow);
          writeCol = 0;
          for (var cell in write) {
            excel.updateCell(sheet, writeCol, writeRow, cell);
            writeCol++;
          }
          writeRow++;
        } else {
          print('skipping row: $row');
        }
      } else {
        print('skipping row: $row');
      }
    }
  }
  final format = DateFormat('yyyy-MM-dd-HH-MM-SS');
  File('out/${format.format(DateTime.now())}.xlsx')
    ..createSync(recursive: true)
    ..writeAsBytesSync(excel.encode());
}

final weekdays = {
  'Lundi': 'MONDAY',
  'Mardi': 'TUESDAY',
  'Mercredi': 'WEDNESDAY',
  'Jeudi': 'THURSDAY',
  'Vendredi': 'FRIDAY',
  'Samedi': 'SATURDAY',
  'Dimanche': 'SUNDAY',
  'lundi': 'MONDAY',
  'mardi': 'TUESDAY',
  'mercredi': 'WEDNESDAY',
  'jeudi': 'THURSDAY',
  'vendredi': 'FRIDAY',
  'samedi': 'SATURDAY',
  'dimanche': 'SUNDAY',
};

double toMicrosoftDateTime(DateTime date) {
  final startOfTime = DateTime(1899, 12, 30);
  final diff = date.difference(startOfTime);
  return diff.inDays + ((diff - Duration(days: diff.inDays)).inSeconds / 86400);
}

DateTime toDartDateTime(double timeStamp) {
  final startOfTime = DateTime(1899, 12, 30);
  return startOfTime.add(Duration(
      days: timeStamp.toInt(), seconds: (timeStamp % 1 * 86400).toInt()));
}

DateTime addTime(DateTime date, String time) {
  var digits = time.split(':');
  return date.add(Duration(
      hours: int.parse(digits[0]),
      minutes: int.parse(digits[1]),
      seconds: int.parse(digits[2])));
}

bool isTimeStamp(Object time) {
  if (time is! String) {
    return false;
  }
  final digits = time.split(':');
  if (digits.length == 3) {
    for (var digit in digits) {
      if (int.tryParse(digit) is! int) {
        return false;
      }
      return true;
    }
  }
  return false;
}
