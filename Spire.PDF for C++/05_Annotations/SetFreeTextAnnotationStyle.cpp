#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_4.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetFreeTextAnnotationStyle.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<RectangleF> rect = new RectangleF(150, 120, 150, 30);
	intrusive_ptr<PdfFreeTextAnnotation> textAnnotation = new PdfFreeTextAnnotation(rect);
	textAnnotation->SetText(L"\nFree Text Annotation Formatting");
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::TimesRoman, 10);
	intrusive_ptr<PdfAnnotationBorder> border = new PdfAnnotationBorder(1.f);
	textAnnotation->SetFont(font);
	textAnnotation->SetBorder(border);
	textAnnotation->SetBorderColor(new PdfRGBColor(Spire::Common::Color::GetPurple()));
	textAnnotation->SetLineEndingStyle(PdfLineEndingStyle::Circle);
	textAnnotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetGreen()));
	textAnnotation->SetOpacity(0.8f);
	page->GetAnnotationsWidget()->Add(textAnnotation);

	rect = new RectangleF(150, 200, 150, 40);
	textAnnotation = new PdfFreeTextAnnotation(rect);
	textAnnotation->SetText(L"\nFree Text Annotation Formatting");
	border = new PdfAnnotationBorder(1.f);
	font = new PdfFont(PdfFontFamily::Helvetica, 10);
	textAnnotation->SetFont(font);
	textAnnotation->SetBorder(border);
	textAnnotation->SetBorderColor(new PdfRGBColor(Spire::Common::Color::GetLightGoldenrodYellow()));
	textAnnotation->SetLineEndingStyle(PdfLineEndingStyle::RClosedArrow);
	textAnnotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetLightPink()));
	textAnnotation->SetOpacity(0.8f);
	page->GetAnnotationsWidget()->Add(textAnnotation);

	rect = new RectangleF(150, 280, 280, 40);
	textAnnotation = new PdfFreeTextAnnotation(rect);
	textAnnotation->SetText(L"\nHow to Set Free Text Annotation Formatting in Pdf file");
	border = new PdfAnnotationBorder(1.f);
	font = new PdfFont(PdfFontFamily::Helvetica, 10);
	textAnnotation->SetFont(font);
	textAnnotation->SetBorder(border);
	textAnnotation->SetBorderColor(new PdfRGBColor(Spire::Common::Color::GetGray()));
	textAnnotation->SetLineEndingStyle(PdfLineEndingStyle::Circle);
	textAnnotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetLightSkyBlue()));
	textAnnotation->SetOpacity(0.8f);
	page->GetAnnotationsWidget()->Add(textAnnotation);

	rect = new RectangleF(150, 360, 200, 40);
	textAnnotation = new PdfFreeTextAnnotation(rect);
	textAnnotation->SetText(L"\nFree Text Annotation Formatting");
	border = new PdfAnnotationBorder(1.f);
	font = new PdfFont(PdfFontFamily::Helvetica, 10);
	textAnnotation->SetFont(font);
	textAnnotation->SetBorder(border);
	textAnnotation->SetBorderColor(new PdfRGBColor(Spire::Common::Color::GetPink()));
	textAnnotation->SetLineEndingStyle(PdfLineEndingStyle::RClosedArrow);
	textAnnotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetLightGreen()));
	textAnnotation->SetOpacity(0.8f);
	page->GetAnnotationsWidget()->Add(textAnnotation);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}

