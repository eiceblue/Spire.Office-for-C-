#include "pch.h"

using namespace std;
using namespace Spire::Pdf;
using namespace Spire::Common;

int main()
{
	wstring outputFile = OUTPUTPATH"InlineImageAndPageNumber.pdf";
	//Load the Pdf from disk
	wstring inputFile = DATAPATH"PDFTemplate_HF.pdf";
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first page
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	wstring text1 = L"Spire.Pdf is a robust component by";
	wstring text2 = L"Technology Co., Ltd.";
	wstring inputFile1 = DATAPATH"Top-logo.png";
	intrusive_ptr<PdfImage> image = PdfImage::FromFile(inputFile1.c_str());

	//Define font and bursh     
	LPCWSTR_S impact = L"Impact";
	intrusive_ptr<PdfTrueTypeFont> font = new PdfTrueTypeFont(impact, 10.f);
	intrusive_ptr<PdfBrush> bursh = PdfBrushes::GetDarkGray();

	//Get the size of text
	intrusive_ptr<SizeF> s1 = font->MeasureString(text1.c_str());
	intrusive_ptr<SizeF> s2 = font->MeasureString(text2.c_str());

	float x = 10;
	float y = 10;

	//Define image size
	intrusive_ptr<SizeF> imgSize = new SizeF(image->GetWidth() / 2, image->GetHeight() / 2);

	//Define rectangle and string format
	intrusive_ptr<SizeF> size = new  SizeF(s1->GetWidth(), imgSize->GetWidth());
	intrusive_ptr<RectangleF> rect1 = new RectangleF(new PointF(x, y), size);
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat(PdfTextAlignment::Left, PdfVerticalAlignment::Middle);

	//Draw the text1
	page->GetCanvas()->DrawString(text1.c_str(), font, bursh, rect1, format);

	//Draw the image
	x += s1->GetWidth();
	page->GetCanvas()->DrawImage(image, new PointF(x, y), imgSize);

	//Draw the text2
	x += imgSize->GetWidth();
	size = new SizeF(s2->GetWidth(), imgSize->GetHeight());
	rect1 = new RectangleF(new PointF(x, y), size);
	page->GetCanvas()->DrawString(text2.c_str(), font, bursh, rect1, format);

	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}

