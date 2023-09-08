#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"PreventOrAllowChangingShape.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Add a rectangle shape to the slide
	intrusive_ptr<IAutoShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(50, 100, 400, 150));

	//Set the shape format
	shape->GetFill()->SetFillType(FillFormatType::None);
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetLightBlue());
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Justify);
	shape->GetTextFrame()->SetText(L"Demo for locking shapes:\n    Green/Black stands for editable.\n    Grey stands for non-editable.");
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Arial Rounded MT Bold"));
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());

	//The changes of selection and rotation are allowed
	shape->GetLocking()->SetRotationProtection(false);
	shape->GetLocking()->SetSelectionProtection(false);
	//The changes of size, position, shape type, aspect ratio, text editing and ajust handles are not allowed 
	shape->GetLocking()->SetResizeProtection(true);
	shape->GetLocking()->SetPositionProtection(true);
	shape->GetLocking()->SetShapeTypeProtection(true);
	shape->GetLocking()->SetAspectRatioProtection(true);
	shape->GetLocking()->SetTextEditingProtection(true);
	shape->GetLocking()->SetAdjustHandlesProtection(true);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

