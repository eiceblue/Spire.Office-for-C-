#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"AutoFitTextOrShape.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()
		->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Set the AutofitType property to Shape
	intrusive_ptr<IAutoShape> textShape2 = ppt->GetSlides()->GetItem(0)->GetShapes()->
		AppendShape(ShapeType::Rectangle, new RectangleF(150, 100, 150, 80));
	textShape2->GetTextFrame()->SetText(L"Resize shape to fit text.");
	textShape2->GetTextFrame()->SetAutofitType(TextAutofitType::Shape);

	//Set the AutofitType property to Normal
	intrusive_ptr<IAutoShape> textShape1 = ppt->GetSlides()->GetItem(0)->GetShapes()
		->AppendShape(ShapeType::Rectangle, new RectangleF(400, 100, 150, 80));
	textShape1->GetTextFrame()->SetText(L"Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape.");
	textShape1->GetTextFrame()->SetAutofitType(TextAutofitType::Normal);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);


}
