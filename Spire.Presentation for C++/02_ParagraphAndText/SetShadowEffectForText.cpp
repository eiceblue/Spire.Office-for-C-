#include "pch.h"

using namespace Spire::Presentation;

int main()
{

	wstring outputFile = OUTPUTPATH"SetShadowEffect.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Set background image
	wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0)->GetLine()->GetFillFormat()->GetSolidFillColor()->SetColor(Color::GetFloralWhite());

	//Get reference of the slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Add a new rectangle shape to the first slide
	intrusive_ptr<IAutoShape> shape = slide->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(120, 100, 450, 200));
	shape->GetFill()->SetFillType(FillFormatType::None);

	//Add the text to the shape and set the font for the text
	shape->AppendTextFrame(L"Text shading on slides");
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Arial Black"));
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetFontHeight(21);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->SetFillType(FillFormatType::Solid);
	shape->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->GetFill()->GetSolidColor()->SetColor(Color::GetBlack());

	////Add inner shadow and set all necessary parameters
	//InnerShadowEffect Shadow = InnerShadowEffect();

	//Add outer shadow and set all necessary parameters
	intrusive_ptr<OuterShadowEffect> Shadow = new OuterShadowEffect();

	Shadow->SetBlurRadius(0);
	Shadow->SetDirection(50);
	Shadow->SetDistance(10);
	Shadow->GetColorFormat()->SetColor(Color::GetLightBlue());

	//shape.TextFrame.TextRange.EffectDag.InnerShadowEffect = Shadow;
	shape->GetTextFrame()->GetTextRange()->GetEffectDag()->SetOuterShadowEffect(Shadow);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

