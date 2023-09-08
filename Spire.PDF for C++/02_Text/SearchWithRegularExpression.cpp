#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SearchReplaceTemplate.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SearchWithRegularExpression.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfTextFindCollection> collection = page->FindText(L"\\d{4}", Find_TextFindParameter::Regex);
	wstring newText = L"New Year";
	intrusive_ptr<PdfBrush> brush = new PdfSolidBrush(new PdfRGBColor(Color::GetDarkBlue()));
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(L"Arial", 7.f, PdfFontStyle::Bold, true);
	intrusive_ptr<PdfStringFormat> centerAlign = new PdfStringFormat(PdfTextAlignment::Center, PdfVerticalAlignment::Middle);
	intrusive_ptr<RectangleF> rec;
	for(intrusive_ptr<PdfTextFind> find : collection->GetFinds())
	{
		// Gets the bound of the found text in page
		rec = find->GetBounds();

		page->GetCanvas()->DrawRectangle(PdfBrushes::GetGreenYellow(), rec);
		// Draws new text as defined font and color
		page->GetCanvas()->DrawString(newText.c_str(), font, brush, rec, centerAlign);

		// This method can directly replace old text with newText,but it just can set the background color, can not set font/forecolor
		// find->ApplyRecoverString(newText.c_str(), Color::GetGray());
	}
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}


