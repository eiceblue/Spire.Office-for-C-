#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"Set3DEffectForText.pptx";

	//Create a new presentation object
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Append a new shape to slide and set the line color and fill type
	intrusive_ptr<IAutoShape> shape = slide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(30, 40, 650, 200));
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetWhite());
	shape->GetFill()->SetFillType(FillFormatType::None);

	//Add text to the shape
	shape->AppendTextFrame(L"This demo shows how to add 3D effect text to Presentation slide");

	//Set the color of text in shape
	intrusive_ptr<TextRange> textRange = shape->GetTextFrame()->GetTextRange();
	textRange->GetFill()->SetFillType(FillFormatType::Solid);
	textRange->GetFill()->GetSolidColor()->SetColor(Color::GetLightBlue());

	//Set the Font of text in shape
	textRange->SetFontHeight(40);
	textRange->SetLatinFont(new TextFont(L"Gulim"));

	//Set 3D effect for text
	shape->GetTextFrame()->GetTextThreeD()->GetShapeThreeD()->SetPresetMaterial(PresetMaterialType::Matte);
	shape->GetTextFrame()->GetTextThreeD()->GetLightRig()->SetPresetType(PresetLightRigType::Sunrise);
	shape->GetTextFrame()->GetTextThreeD()->GetShapeThreeD()->GetTopBevel()->SetPresetType(BevelPresetType::Circle);
	shape->GetTextFrame()->GetTextThreeD()->GetShapeThreeD()->GetContourColor()->SetColor(Color::GetGreen());
	shape->GetTextFrame()->GetTextThreeD()->GetShapeThreeD()->SetContourWidth(3);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

