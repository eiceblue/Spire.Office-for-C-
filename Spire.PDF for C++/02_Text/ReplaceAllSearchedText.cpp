#include "../pch.h"

using namespace Spire::Pdf;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path+L"SearchReplaceTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"ReplaceAllSearchedText.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfTextFindCollection> collection = page->FindText(L"Spire.PDF for .NET", Find_TextFindParameter::IgnoreCase);
	wstring newText = L"E-iceblue Spire.PDF";
	intrusive_ptr<PdfBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetDarkBlue()));
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
	intrusive_ptr<RectangleF> rec;
	for(intrusive_ptr<PdfTextFind> find : collection->GetFinds())
	{
		rec = find->GetBounds();
		page->GetCanvas()->DrawRectangle(PdfBrushes::GetWhite(), rec);
		page->GetCanvas()->DrawString(newText.c_str(), font, brush, rec);
	}
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

