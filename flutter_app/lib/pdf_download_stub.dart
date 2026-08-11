// Stub untuk platform Web — PDF download tidak tersedia via channel
// di Web, download sudah ditangani langsung oleh JavaScript di browser
Future<String> savePdfNative(String base64Str, String fileName) async {
  return '✅ PDF siap diunduh melalui browser';
}
