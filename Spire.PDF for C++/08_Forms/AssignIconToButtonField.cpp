#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring input_path = DATAPATH;
	wstring input1 = input_path + L"E-iceblueLogo.png";
	wstring input2 = input_path + L"PdfImage.png";
	wstring input3 = input_path + L"E-logo.png";

	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AssignIconToButtonField.pdf";

	//Create a PDF document       
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();

	//Create a Button
	intrusive_ptr<PdfButtonField> btn = new PdfButtonField(page, L"button1");
	btn->SetBounds(new RectangleF(0, 50, 120, 100));
	btn->SetHighlightMode(PdfHighlightMode::Push);
	btn->SetLayoutMode(PdfButtonLayoutMode::CaptionOverlayIcon);

	//Set text and icon for Normal appearance
	btn->SetText(L"Normal Text");
	btn->SetIcon(PdfImage::FromFile(input1.c_str()));

	//Set text and icon for Click appearance (Can only be set when highlight mode is Push)
	btn->SetAlternateText(L"Alternate Text");
	btn->SetAlternateIcon(PdfImage::FromFile(input2.c_str()));

	//Set text and icon for Rollover appearance (Can only be set when highlight mode is Push)
	btn->SetRolloverText(L"Rollover Text");
	btn->SetRolloverIcon(PdfImage::FromFile(input3.c_str()));

	//Set icon layout
	btn->GetIconLayout()->SetSpaces({ 0.5f, 0.5f });
	btn->GetIconLayout()->SetScaleMode(PdfButtonIconScaleMode::Proportional);
	btn->GetIconLayout()->SetScaleReason(PdfButtonIconScaleReason::Always);
	btn->GetIconLayout()->SetIsFitBounds(false);

	//Add the button to the document
	doc->GetForm()->GetFields()->Add(btn);
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}