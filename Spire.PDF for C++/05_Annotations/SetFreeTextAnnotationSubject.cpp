#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_4.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetFreeTextAnnotationSubject.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<RectangleF> rect = new RectangleF(150, 120, 150, 30);
	intrusive_ptr<PdfFreeTextAnnotation> textAnnotation = new PdfFreeTextAnnotation(rect);
	textAnnotation->SetText(L"\nSet free text annotation subject");
	textAnnotation->SetSubject(L"SubjectTest");
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::TimesRoman, 10);
	intrusive_ptr<PdfAnnotationBorder> border = new PdfAnnotationBorder(1.f);
	textAnnotation->SetFont(font);
	textAnnotation->SetBorder(border);
	textAnnotation->SetBorderColor(new PdfRGBColor(Spire::Common::Color::GetPurple()));
	textAnnotation->SetLineEndingStyle(PdfLineEndingStyle::Circle);
	textAnnotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetGreen()));
	textAnnotation->SetOpacity(0.8f);
	page->GetAnnotationsWidget()->Add(textAnnotation);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}


