
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile =DATAPATH"SetBackground.pptx";
	wstring outputFile = OUTPUTPATH"SetBackground.pptx";

	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());
	//Set the background of the first slide to Gradient color
	presentation->GetSlides()->GetItem(0)->GetSlideBackground()->SetType(BackgroundType::Custom);
	presentation->GetSlides()->GetItem(0)->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Gradient);
	presentation->GetSlides()->GetItem(0)->GetSlideBackground()->GetFill()->GetGradient()->SetGradientShape(GradientShapeType::Linear);
	presentation->GetSlides()->GetItem(0)->GetSlideBackground()->GetFill()->GetGradient()->SetGradientStyle(GradientStyle::FromCorner1);
	presentation->GetSlides()->GetItem(0)->GetSlideBackground()->GetFill()->GetGradient()->GetGradientStops()->Append(1.0f, KnownColors::SkyBlue);
	presentation->GetSlides()->GetItem(0)->GetSlideBackground()->GetFill()->GetGradient()->GetGradientStops()->Append(0.0f, KnownColors::White);

	//Set the background of the second slide to Solid color
	presentation->GetSlides()->GetItem(1)->GetSlideBackground()->SetType(BackgroundType::Custom);
	presentation->GetSlides()->GetItem(1)->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Solid);
	presentation->GetSlides()->GetItem(1)->GetSlideBackground()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::SkyBlue);

	presentation->GetSlides()->Append();
	//Set the background of the third slide to picture
	std::wstring ImageFile = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, presentation->GetSlideSize()->GetSize()->GetWidth(), presentation->GetSlideSize()->GetSize()->GetHeight());
	presentation->GetSlides()->GetItem(2)->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Picture);

	intrusive_ptr<IEmbedImage> image = presentation->GetSlides()->GetItem(2)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, ImageFile.c_str(), rect);
	//presentation->GetSlides()->GetItem(2)->GetSlideBackground()->GetFill()->GetPictureFill()->GetPicture()->SetEmbedImage(Object::Dynamic_cast<IImageData>>(image));

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
