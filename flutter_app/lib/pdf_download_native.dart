import 'dart:convert';
import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:path_provider/path_provider.dart';

/// Simpan PDF ke penyimpanan lokal (Android & iOS) dengan penanganan nama unik
Future<String> savePdfNative(String base64Str, String fileName) async {
  try {
    final Uint8List bytes = base64Decode(base64Str);

    Directory dir;

    if (Platform.isAndroid) {
      try {
        final Directory baseDownloadsDir =
            Directory('/storage/emulated/0/Download');
        Directory targetDir;
        if (await baseDownloadsDir.exists()) {
          targetDir = Directory('${baseDownloadsDir.path}/Warnasari Data');
        } else {
          final extDir = await getExternalStorageDirectory() ??
              await getApplicationDocumentsDirectory();
          targetDir = Directory('${extDir.path}/Warnasari Data');
        }
        if (!await targetDir.exists()) {
          await targetDir.create(recursive: true);
        }
        dir = targetDir;
      } catch (_) {
        dir = await getApplicationDocumentsDirectory();
      }
    } else {
      dir = await getApplicationDocumentsDirectory();
    }

    // Olah nama file agar selalu unik (tidak menimpa file lama)
    String trimmedName = fileName.trim();
    String nameWithoutExt = trimmedName;
    String ext = '.pdf';
    int lastDot = trimmedName.lastIndexOf('.');
    if (lastDot > 0) {
      nameWithoutExt = trimmedName.substring(0, lastDot);
      ext = trimmedName.substring(lastDot);
    }

    // Tambahkan timestamp jam-menit-detik untuk memastikan file unik
    final now = DateTime.now();
    final timeStamp =
        '${now.hour.toString().padLeft(2, '0')}${now.minute.toString().padLeft(2, '0')}${now.second.toString().padLeft(2, '0')}';
    
    String finalFileName = '${nameWithoutExt}_$timeStamp$ext';
    String filePath = '${dir.path}/$finalFileName';

    // Jika karena suatu hal file tersebut masih ada, tambahkan penomoran counter
    int counter = 1;
    while (await File(filePath).exists()) {
      finalFileName = '${nameWithoutExt}_${timeStamp}_($counter)$ext';
      filePath = '${dir.path}/$finalFileName';
      counter++;
    }

    await File(filePath).writeAsBytes(bytes, flush: true);

    String fileTypeLabel = ext.toLowerCase() == '.json'
        ? 'Data JSON'
        : (ext.toLowerCase() == '.pdf' ? 'PDF' : 'File');

    if (Platform.isAndroid) {
      return '✅ $fileTypeLabel disimpan!\nBuka Downloads → Warnasari Data → $finalFileName';
    } else {
      return '✅ $fileTypeLabel disimpan!\nBuka Files app → $finalFileName';
    }
  } catch (e) {
    debugPrint('savePdfNative error: $e');
    return '⚠️ Gagal menyimpan file: $e';
  }
}
