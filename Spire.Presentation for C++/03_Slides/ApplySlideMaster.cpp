#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"InputTemplate.pptx";
	wstring outputFile = OUTPUTPATH"ApplySlideMaster.pptx";

	//Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide master from the presentation
	intrusive_ptr<IMasterSlide> masterSlide = ppt->GetMasters()->GetItem(0);

	//Customize the background of the slide master
	wstring backgroundPic = DATAPATH"bg.png";
	intrusive_ptr<RectangleF> rect = new RectangleF(0, 0, ppt->GetSlideSize()->GetSize()->GetWidth(), ppt->GetSlideSize()->GetSize()->GetHeight());
	masterSlide->GetSlideBackground()->GetFill()->SetFillType(FillFormatType::Picture);
	intrusive_ptr<IEmbedImage> image = masterSlide->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, backgroundPic.c_str(), rect);
	masterSlide->GetSlideBackground()->GetFill()->GetPictureFill()->GetPicture()->SetEmbedImage(image->GetPictureFill()->GetPicture()->GetEmbedImage());

	//Change the color scheme
	masterSlide->GetTheme()->GetColorScheme()->GetAccent1()->SetColor(Color::GetRed());
	masterSlide->GetTheme()->GetColorScheme()->GetAccent2()->SetColor(Color::GetRosyBrown());
	masterSlide->GetTheme()->GetColorScheme()->GetAccent3()->SetColor(Color::GetIvory());
	masterSlide->GetTheme()->GetColorScheme()->GetAccent4()->SetColor(Color::GetLavender());
	masterSlide->GetTheme()->GetColorScheme()->GetAccent5()->SetColor(Color::GetBlack());

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
