#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"Set3DEffectForShape.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Add shape1 and fill it with color
	intrusive_ptr<IAutoShape> shape1 = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::RoundCornerRectangle, new RectangleF(150, 150, 150, 150));
	shape1->GetFill()->SetFillType(FillFormatType::Solid);
	shape1->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::SkyBlue);
	//Initialize a new instance of the 3-D class for shape1 and set its properties
	intrusive_ptr<ShapeThreeD> effect1 = shape1->GetThreeD()->GetShapeThreeD();
	effect1->SetPresetMaterial(PresetMaterialType::Powder);
	effect1->GetTopBevel()->SetPresetType(BevelPresetType::ArtDeco);
	effect1->GetTopBevel()->SetHeight(4);
	effect1->GetTopBevel()->SetWidth(12);
	effect1->SetBevelColorMode(BevelColorType::Contour);
	effect1->GetContourColor()->SetKnownColor(KnownColors::LightBlue);
	effect1->SetContourWidth(3.5);

	//Add shape2 and fill it with color
	intrusive_ptr<IAutoShape> shape2 = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Pentagon, new RectangleF(400, 150, 150, 150));
	shape2->GetFill()->SetFillType(FillFormatType::Solid);
	shape2->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::LightGreen);
	//Initialize a new instance of the 3-D class for shape2 and set its properties
	intrusive_ptr<ShapeThreeD> effect2 = shape2->GetThreeD()->GetShapeThreeD();
	effect2->SetPresetMaterial(PresetMaterialType::SoftEdge);
	effect2->GetTopBevel()->SetPresetType(BevelPresetType::SoftRound);
	effect2->GetTopBevel()->SetHeight(12);
	effect2->GetTopBevel()->SetWidth(12);
	effect2->SetBevelColorMode(BevelColorType::Contour);
	effect2->GetContourColor()->SetKnownColor(KnownColors::LawnGreen);
	effect2->SetContourWidth(5);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}

