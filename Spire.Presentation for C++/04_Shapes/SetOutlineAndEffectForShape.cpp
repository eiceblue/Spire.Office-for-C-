#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"SetOutlineAndEffectForShape.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Get the first slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Set background Image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	slide->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	slide->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Draw a Rectangle shape
	intrusive_ptr<IAutoShape> shape = slide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(150, 180, 100, 50));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetSkyBlue());
	//Set outline color
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetRed());
	//Set shadow effect
	intrusive_ptr<PresetShadow> shadow = new PresetShadow();
	shadow->GetColorFormat()->SetColor(Color::GetLightSkyBlue());
	shadow->SetPreset(PresetShadowValue::FrontRightPerspective);
	shadow->SetDistance(10.0);
	shadow->SetDirection(225.0f);
	shape->GetEffectDag()->SetPresetShadowEffect(shadow);

	//Draw a Ellipse shape
	shape = slide->GetShapes()->AppendShape(ShapeType::Ellipse, new RectangleF(400, 150, 100, 100));
	shape->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetFill()->GetSolidColor()->SetColor(Color::GetSkyBlue());
	//Set outline color
	shape->GetShapeStyle()->GetLineColor()->SetColor(Color::GetYellow());
	//Set shadow effect
	intrusive_ptr<GlowEffect> glow = new GlowEffect();
	glow->GetColorFormat()->SetColor(Color::GetLightPink());
	glow->SetRadius(20.0);
	shape->GetEffectDag()->SetGlowEffect(glow);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

