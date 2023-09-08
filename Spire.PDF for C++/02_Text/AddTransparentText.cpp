#include "../pch.h"

using namespace Spire::Pdf;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path +L"AddtransparentText.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<SizeF> size = new SizeF(PdfPageSize::A4());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add(size, new PdfMargins(0));
	page->GetCanvas()->Save();
	float alpha = 0.25f;
	page->GetCanvas()->SetTransparency(alpha, alpha, PdfBlendMode::Normal);
	intrusive_ptr<RectangleF> rect = new RectangleF(50, 50, 450, page->GetSize()->GetHeight());
	wstring text = L"Spire.PDF for Cpp,a professional PDF library applied to creating, writing, editing, handling and reading PDF files without any external dependencies within .NET( C#, VB.NET, ASP.NET, .NET Core) application.";
	text += L"\n\n\n\n\n";
	text += L"Spire.PDF for Cpp,a PDF C++ API that enables developers to read, write, convert and print PDF documents in C++ applications without using Adobe Acrobat.";
	LPCWSTR_S str = text.c_str();
	intrusive_ptr<PdfRGBColor> color = new PdfRGBColor(Color::FromArgb(30, 0, 255, 0));
	intrusive_ptr<PdfSolidBrush> brush = new PdfSolidBrush(color);
	page->GetCanvas()->DrawString(str, new PdfFont(PdfFontFamily::Helvetica, 14.f), brush, rect);
	page->GetCanvas()->Restore();
	doc->SaveToFile((outputFile).c_str());
	doc->Close();
}

