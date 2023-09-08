#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreatePdfLineAnnotation.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	int* linePoints = new int[4] { 100, 650, 180, 650 };
	vector<int> vec;
	vec.insert(vec.begin(), linePoints, linePoints + 4);
	intrusive_ptr<PdfLineAnnotation> lineAnnotation = new PdfLineAnnotation(vec, L"This is the first line annotation");
	lineAnnotation->GetlineBorder()->SetBorderStyle(PdfBorderStyle::Solid);
	lineAnnotation->GetlineBorder()->SetBorderWidth(1);
	lineAnnotation->SetLineIntent(PdfLineIntent::LineDimension);
	lineAnnotation->SetBeginLineStyle(PdfLineEndingStyle::Butt);
	lineAnnotation->SetEndLineStyle(PdfLineEndingStyle::Diamond);
	lineAnnotation->SetFlags(PdfAnnotationFlags::Default);
	lineAnnotation->SetInnerLineColor(new PdfRGBColor(Spire::Common::Color::GetGreen()));
	lineAnnotation->SetBackColor(new PdfRGBColor(Spire::Common::Color::GetGreen()));
	lineAnnotation->SetLeaderLineExt(0);
	lineAnnotation->SetLeaderLine(0);
	page->GetAnnotationsWidget()->Add(lineAnnotation);

	linePoints = new int[4] { 100, 550, 280, 550 };
	vector<int> vec1(linePoints, linePoints + sizeof(linePoints) / sizeof(int));
	lineAnnotation = new PdfLineAnnotation(vec1, L"This is the second line annotation");
	lineAnnotation->GetlineBorder()->SetBorderStyle(PdfBorderStyle::Underline);
	lineAnnotation->GetlineBorder()->SetBorderWidth(2);
	lineAnnotation->SetLineIntent(PdfLineIntent::LineArrow);
	lineAnnotation->SetBeginLineStyle(PdfLineEndingStyle::Circle);
	lineAnnotation->SetEndLineStyle(PdfLineEndingStyle::Diamond);
	lineAnnotation->SetFlags(PdfAnnotationFlags::Default);
	lineAnnotation->SetInnerLineColor(new PdfRGBColor(Spire::Common::Color::GetPink()));
	lineAnnotation->SetBackColor(new PdfRGBColor(Spire::Common::Color::GetPink()));
	lineAnnotation->SetLeaderLineExt(0);
	lineAnnotation->SetLeaderLine(0);
	page->GetAnnotationsWidget()->Add(lineAnnotation);

	linePoints = new int[4] { 100, 450, 280, 450 };
	vector<int> vec2(linePoints, linePoints + sizeof(linePoints) / sizeof(int));
	lineAnnotation = new PdfLineAnnotation(vec2, L"This is the third line annotation");
	lineAnnotation->GetlineBorder()->SetBorderStyle(PdfBorderStyle::Beveled);
	lineAnnotation->GetlineBorder()->SetBorderWidth(2);
	lineAnnotation->SetLineIntent(PdfLineIntent::LineDimension);
	lineAnnotation->SetBeginLineStyle(PdfLineEndingStyle::None);
	lineAnnotation->SetEndLineStyle(PdfLineEndingStyle::None);
	lineAnnotation->SetFlags(PdfAnnotationFlags::Default);
	lineAnnotation->SetInnerLineColor(new PdfRGBColor(Spire::Common::Color::GetBlue()));
	lineAnnotation->SetBackColor(new PdfRGBColor(Spire::Common::Color::GetBlue()));
	lineAnnotation->SetLeaderLineExt(1);
	lineAnnotation->SetLeaderLine(1);
	page->GetAnnotationsWidget()->Add(lineAnnotation);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}


