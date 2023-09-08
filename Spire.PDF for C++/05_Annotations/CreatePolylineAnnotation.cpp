#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreatePolylineAnnotation.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	vector<intrusive_ptr<PointF>> vec{ new PointF(0, 60), new PointF(30, 45), new PointF(60, 90), new PointF(90, 80) };
	intrusive_ptr<PdfPolyLineAnnotation> polyline = new PdfPolyLineAnnotation(page, vec);
	polyline->SetColor(new PdfRGBColor(Spire::Common::Color::GetPaleVioletRed()));
	polyline->SetText(L"This is a polygon annotation");
	polyline->SetAuthor(L"E-ICEBLUE");
	polyline->SetSubject(L"polygon annotation demo");
	polyline->SetBorder(new PdfAnnotationBorder(1.f));
	polyline->SetModifiedDate(DateTime::GetNow());
	page->GetAnnotationsWidget()->Add(polyline);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

