import 'package:flutter_test/flutter_test.dart';
import 'package:datadusun1/main.dart';

void main() {
  testWidgets('App smoke test', (WidgetTester tester) async {
    await tester.pumpWidget(const WarnasariDataApp());
    expect(find.byType(WarnasariDataApp), findsOneWidget);
  });
}
