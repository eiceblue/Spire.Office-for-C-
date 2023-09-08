#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetLineBreak.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<SizeF> size = new SizeF(PdfPageSize::A4());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add(size, new PdfMargins(40));
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetBlack()));
	wstring text = L"Spire.PDF for Cpp";
	text += L"A professional PDF library applied to";
	text += L" creating, writing, editing, handling and reading PDF files";
	text += L" without any external dependencies within";
	text += L"C++ application.";
	text += L"\n\rSpire.PDF for Java\n";
	text += L"A PDF Java API that enables developers to read, ";
	text += L"write, convert and print PDF documents";
	text += L"in Java applications without using Adobe Acrobat.";
	text += L"\n\r";
	text += L"Welcome to evaluate Spire.PDF!";
	intrusive_ptr<RectangleF> rect = new RectangleF(50, 50, page->GetSize()->GetWidth() - 150, page->GetSize()->GetHeight());
	page->GetCanvas()->DrawString(text.c_str(), new PdfFont(PdfFontFamily::Helvetica, 13.f), brush, rect);
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}


