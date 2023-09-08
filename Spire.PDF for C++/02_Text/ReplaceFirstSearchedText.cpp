#include "../pch.h"

using namespace Spire::Pdf;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile =input_path+L"SearchReplaceTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"ReplaceFirstSearchedText.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	
	intrusive_ptr<PdfTextFindCollection> collection = page->FindText(L"Spire.PDF for C++", Find_TextFindParameter::IgnoreCase);
	wstring newText = L"Spire.PDF API";
	intrusive_ptr<PdfTextFind> find = collection->GetFinds()[0];
	intrusive_ptr<PdfBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetDarkBlue()));
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 15.f, PdfFontStyle::Bold, true);
	intrusive_ptr<RectangleF> rec = find->GetBounds();
	page->GetCanvas()->DrawRectangle(PdfBrushes::GetWhite(), rec);
	page->GetCanvas()->DrawString(newText.c_str(), font, brush, rec);
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
