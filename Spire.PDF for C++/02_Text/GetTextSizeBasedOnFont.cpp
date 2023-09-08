#include "../pch.h"

using namespace Spire::Pdf;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"GetTextSizeBasedOnFont.txt";

	wstring text = L"Spire.PDF for Cpp";
	intrusive_ptr<PdfFont> font1 = new PdfFont(PdfFontFamily::TimesRoman, 12.f);
	intrusive_ptr<SizeF> sizeF1 = font1->MeasureString(text.c_str());
	wstring str1 = L"1. The width is: ";
	str1 += to_wstring(sizeF1->GetWidth());
	str1 += L", the height is: ";
	str1 += to_wstring(sizeF1->GetHeight());
	intrusive_ptr<PdfTrueTypeFont> font2 = new PdfTrueTypeFont(L"Arial", 12.f, PdfFontStyle::Regular, true);
	intrusive_ptr<SizeF> sizeF2 = font2->MeasureString(text.c_str());
	wstring str2 = L"2. The width is: ";
	str2 += to_wstring(sizeF2->GetWidth());
	str2 += L", the height is: ";
	str2 += to_wstring(sizeF2->GetHeight());
	wofstream os;
	os.open(outputFile, ios::trunc);
	os << str1 << "\n";
	os << str2 << "\n";
	os.close();
}

