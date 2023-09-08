#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

inline PdfAnnotationFlags operator|(PdfAnnotationFlags a, PdfAnnotationFlags b) {
				return static_cast<PdfAnnotationFlags>(static_cast<int>(a) | static_cast<int>(b));
			}
			
int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_4.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InvisibleFreeTextAnnotation.Pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<RectangleF> rect = new RectangleF(100, 120, 150, 30);
	intrusive_ptr<PdfFreeTextAnnotation> FreetextAnnotation = new PdfFreeTextAnnotation(rect);
	FreetextAnnotation->SetText(L"Invisible Free Text Annotation");
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::TimesRoman, 10);
	intrusive_ptr<PdfAnnotationBorder> border = new PdfAnnotationBorder(1.f);
	FreetextAnnotation->SetFont(font);
	FreetextAnnotation->SetBorder(border);
	FreetextAnnotation->SetBorderColor(new PdfRGBColor(Spire::Common::Color::GetPurple()));
	FreetextAnnotation->SetLineEndingStyle(PdfLineEndingStyle::Circle);
	FreetextAnnotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetGreen()));
	FreetextAnnotation->SetOpacity(0.8f);
	FreetextAnnotation->SetFlags(PdfAnnotationFlags::Print | PdfAnnotationFlags::NoView);
	page->GetAnnotationsWidget()->Add(FreetextAnnotation);

	rect = new RectangleF(100, 180, 150, 30);
	FreetextAnnotation = new PdfFreeTextAnnotation(rect);
	FreetextAnnotation->SetText(L"Show Free Text Annotation");
	FreetextAnnotation->SetFont(font);
	FreetextAnnotation->SetBorder(border);
	FreetextAnnotation->SetBorderColor(new PdfRGBColor(Spire::Common::Color::GetLightPink()));
	FreetextAnnotation->SetLineEndingStyle(PdfLineEndingStyle::Circle);
	FreetextAnnotation->SetColor(new PdfRGBColor(Spire::Common::Color::GetLightGreen()));
	FreetextAnnotation->SetOpacity(0.8f);
	page->GetAnnotationsWidget()->Add(FreetextAnnotation);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}


