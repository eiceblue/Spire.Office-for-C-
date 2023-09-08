#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"InsertVideo.pptx";
	wstring outputFile = OUTPUTPATH"InsertVideo.pptx";
	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Add title
	intrusive_ptr<RectangleF> rec_title = new RectangleF(50, 280, 160, 50);

	intrusive_ptr<IAutoShape> shape_title = presentation->GetSlides()->GetItem(0)->GetShapes()
		->AppendShape(ShapeType::Rectangle, rec_title);
	shape_title->GetShapeStyle()->GetLineColor()->SetColor(Color::GetTransparent());

	shape_title->GetFill()->SetFillType(Spire::Presentation::FillFormatType::None);
	intrusive_ptr<TextParagraph> para_title = new TextParagraph();
	std::wstring name = L"Video:";
	std::wstring fontName = L"Myriad Pro Light";
	para_title->SetText(name.c_str());
	para_title->SetAlignment(TextAlignmentType::Center);
	para_title->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(fontName.c_str()));
	para_title->GetTextRanges()->GetItem(0)->SetFontHeight(32);
	para_title->GetTextRanges()->GetItem(0)->SetIsBold(TriState::True);
	para_title->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(Spire::Presentation::FillFormatType::Solid);
	para_title->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::FromArgb(68, 68, 68));
	shape_title->GetTextFrame()->GetParagraphs()->Append(para_title);

	//Insert audio into the document
	intrusive_ptr<RectangleF> audioRect = new RectangleF(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 125, 240, 150, 150);

	std::wstring inputFile1 = DATAPATH"Video.mp4";
	std::wstring inputFile2 = DATAPATH"Video.png";

	intrusive_ptr<IVideo> video = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendVideoMedia(inputFile1.c_str(), audioRect);
	video->GetPictureFill()->GetPicture()->SetUrl(inputFile2.c_str());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

}