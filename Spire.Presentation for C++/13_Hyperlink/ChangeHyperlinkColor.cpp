#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"SmartArtNode.pptx";
	wstring outputFile = OUTPUTPATH"ChangeHyperlinkColor.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(0);

	//Get the theme of the slide
	intrusive_ptr<Theme> theme = slide->GetTheme();

	//Change the color of hyperlink to red
	theme->GetColorScheme()->GetHyperlinkColor()->SetColor(Color::GetRed());

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
