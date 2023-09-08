#include "../pch.h"

using namespace Spire::Pdf;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path+L"ReplaceFont.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"ReplaceFont.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfTrueTypeFont> newFont = new PdfTrueTypeFont(L"Arial", 13.f, PdfFontStyle::Regular, true);
	for(intrusive_ptr<PdfUsedFont> font : doc->GetUsedFonts())
	{
		//Replace the font with new font
		font->Replace(newFont);
	}
	doc->SaveToFile(outputFile.c_str());
	doc->Close();

}

