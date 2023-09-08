#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Pdf_3.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreatePdfLinkAnnotation.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 40, 250, 35);

	intrusive_ptr<PdfFileLinkAnnotation> link = new PdfFileLinkAnnotation(rect, inputFile.c_str());
	page->GetAnnotationsWidget()->Add(link);
	intrusive_ptr<PdfFreeTextAnnotation> text = new PdfFreeTextAnnotation(rect);
	text->SetText(L"Click here! This is a link annotation.");
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 15);
	text->SetFont(font);
	page->GetAnnotationsWidget()->Add(text);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}


