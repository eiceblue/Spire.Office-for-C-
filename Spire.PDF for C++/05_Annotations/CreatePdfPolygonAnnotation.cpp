#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreatePdfPolygonAnnotation.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	vector<intrusive_ptr<PointF>> vec{ new PointF(0, 30), new PointF(30, 15), new PointF(60, 30),
	new PointF(45, 50), new PointF(15, 50), new PointF(0, 30) };
	intrusive_ptr<PdfPolygonAnnotation> polygon = new PdfPolygonAnnotation(page, vec);
	polygon->SetColor(new PdfRGBColor(Spire::Common::Color::GetPaleVioletRed()));
	polygon->SetText(L"This is a polygon annotation");
	polygon->SetAuthor(L"E-ICEBLUE");
	polygon->SetSubject(L"polygon annotation demo");
	polygon->SetBorderEffect(PdfBorderEffect::BigCloud);
	polygon->SetModifiedDate(DateTime::GetNow());
	page->GetAnnotationsWidget()->Add(polygon);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

