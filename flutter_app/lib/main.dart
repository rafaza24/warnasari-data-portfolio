import 'dart:convert';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:webview_flutter/webview_flutter.dart';

// Import platform-specific hanya saat bukan Web
import 'pdf_download_stub.dart'
    if (dart.library.io) 'pdf_download_native.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const WarnasariDataApp());
}

class WarnasariDataApp extends StatelessWidget {
  const WarnasariDataApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Warnasari Data',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        brightness: Brightness.dark,
        scaffoldBackgroundColor: const Color(0xFF0B0F19),
        colorScheme: const ColorScheme.dark(
          primary: Color(0xFF6366F1),
          secondary: Color(0xFF818CF8),
          surface: Color(0xFF151C2C),
        ),
        useMaterial3: true,
      ),
      home: const WebAppScreen(),
    );
  }
}

class WebAppScreen extends StatefulWidget {
  const WebAppScreen({super.key});

  @override
  State<WebAppScreen> createState() => _WebAppScreenState();
}

class _WebAppScreenState extends State<WebAppScreen> {
  late final WebViewController _controller;
  bool _isLoading = true;
  double _progress = 0;
  final String _initialUrl =
      'https://script.google.com/macros/s/YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_ID/exec';

  // ─── Terima base64 PDF dari WebView ─────────────────────────────────────────
  Future<void> _handlePdfDownload(String message) async {
    try {
      final Map<String, dynamic> data = json.decode(message);
      final String base64Str = data['base64'] as String;
      final String fileName = (data['name'] as String?) ?? 'Laporan.pdf';

      // Panggil fungsi platform-specific (native = simpan file, web = tidak digunakan)
      final String result = await savePdfNative(base64Str, fileName);
      if (mounted) _showSnackBar(result);
    } catch (e) {
      debugPrint('PDF download error: $e');
      if (mounted) _showSnackBar('⚠️ Gagal menyimpan PDF: $e');
    }
  }

  void _showSnackBar(String msg, {Duration? duration}) {
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Text(msg, style: const TextStyle(fontSize: 12.5)),
        backgroundColor: const Color(0xFF1E2A3A),
        behavior: SnackBarBehavior.floating,
        duration: duration ?? const Duration(seconds: 5),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
      ),
    );
  }

  @override
  void initState() {
    super.initState();
    final controller = WebViewController();

    if (!kIsWeb) {
      controller
        ..setJavaScriptMode(JavaScriptMode.unrestricted)
        ..setBackgroundColor(const Color(0xFF0B0F19))
        ..addJavaScriptChannel(
          'FlutterDownload',
          onMessageReceived: (JavaScriptMessage msg) {
            _handlePdfDownload(msg.message);
          },
        )
        ..setNavigationDelegate(
          NavigationDelegate(
            onProgress: (int progress) {
              if (mounted) setState(() => _progress = progress / 100.0);
            },
            onPageStarted: (String url) {
              if (mounted) setState(() => _isLoading = true);
            },
            onPageFinished: (String url) {
              if (mounted) setState(() => _isLoading = false);
            },
            onWebResourceError: (WebResourceError error) {
              debugPrint('WebResourceError: ${error.description}');
            },
            onNavigationRequest: (NavigationRequest request) {
              return NavigationDecision.navigate;
            },
          ),
        );
    } else {
      // Pada platform Web, langsung set loading false
      WidgetsBinding.instance.addPostFrameCallback((_) {
        if (mounted) setState(() => _isLoading = false);
      });
    }

    controller.loadRequest(Uri.parse(_initialUrl));
    _controller = controller;
  }

  Future<void> _handlePopScope(bool didPop) async {
    if (didPop) return;
    final nav = Navigator.of(context);
    if (!kIsWeb && await _controller.canGoBack()) {
      await _controller.goBack();
    } else {
      if (!mounted) return;
      final shouldExit = await showDialog<bool>(
        context: context,
        builder: (dialogContext) => AlertDialog(
          backgroundColor: const Color(0xFF151C2C),
          title: const Text('Keluar Aplikasi?'),
          content: const Text(
              'Apakah Anda yakin ingin keluar dari aplikasi Warnasari Data?'),
          actions: [
            TextButton(
              onPressed: () => Navigator.of(dialogContext).pop(false),
              child: const Text('Batal',
                  style: TextStyle(color: Color(0xFF818CF8))),
            ),
            ElevatedButton(
              style: ElevatedButton.styleFrom(
                  backgroundColor: const Color(0xFFEF4444)),
              onPressed: () => Navigator.of(dialogContext).pop(true),
              child:
                  const Text('Keluar', style: TextStyle(color: Colors.white)),
            ),
          ],
        ),
      );
      if (shouldExit == true) nav.pop();
    }
  }

  @override
  Widget build(BuildContext context) {
    return PopScope(
      canPop: false,
      onPopInvokedWithResult: (didPop, result) => _handlePopScope(didPop),
      child: Scaffold(
        body: SafeArea(
          child: Column(
            children: [
              if (_isLoading)
                LinearProgressIndicator(
                  value: _progress > 0 && _progress < 1.0 ? _progress : null,
                  backgroundColor: const Color(0xFF0F1626),
                  color: const Color(0xFF6366F1),
                  minHeight: 3,
                ),
              Expanded(
                child: RefreshIndicator(
                  backgroundColor: const Color(0xFF151C2C),
                  color: const Color(0xFF6366F1),
                  onRefresh: () async {
                    if (!kIsWeb) {
                      await _controller.reload();
                    } else {
                      await _controller.loadRequest(Uri.parse(_initialUrl));
                    }
                  },
                  child: WebViewWidget(controller: _controller),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
