#include "../pch.h"

using namespace std;
using namespace Spire::Pdf;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ReplaceImage.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceImageWithText.pdf";

	//Create a pdf document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile(inputFile.c_str());
	//Get the first page.
	intrusive_ptr<PdfPageBase> page = pdf->GetPages()->GetItem(0);
	//Get images of the first page.
	std::vector<intrusive_ptr<PdfImageInfo>> imageInfo = page->GetImagesInfo();

	intrusive_ptr<Stream> imgStream = imageInfo[0]->GetImage();
	intrusive_ptr<PdfImage> image = new PdfImage();
	image = image->FromStream(imgStream);
	float widthInPixel = image->GetWidth();
	float heightInPixel = image->GetHeight();
	//Convert unit from Pixel to Point
	intrusive_ptr<PdfUnitConvertor> convertor = new PdfUnitConvertor();
	float width = convertor->ConvertFromPixels(widthInPixel, PdfGraphicsUnit::Point);
	float height = convertor->ConvertFromPixels(heightInPixel, PdfGraphicsUnit::Point);
	//Get location of image 
	float xPos = imageInfo[0]->GetBounds()->GetX();
	float yPos = imageInfo[0]->GetBounds()->GetY();

	//Remove the image
	page->DeleteImage(0);

	intrusive_ptr<RectangleF> rect = new RectangleF(new PointF(xPos, yPos), new SizeF(width, height));
	
	intrusive_ptr<PdfStringFormat> format = new PdfStringFormat();
	format->SetAlignment(PdfTextAlignment::Center);
	format->SetLineAlignment(PdfVerticalAlignment::Middle);
	intrusive_ptr<PdfFont> font = new PdfFont(PdfFontFamily::Helvetica, 18.0f);
	wstring s = L"ReplacedText";
	page->GetCanvas()->DrawString(s.c_str(), font, PdfBrushes::GetPurple(), rect, format);
	//Save the document
	pdf->SaveToFile(outputFile.c_str());

	pdf->Close();
}











