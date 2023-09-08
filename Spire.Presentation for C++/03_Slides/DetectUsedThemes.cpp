#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Themes.pptx";
	wstring outputFile = OUTPUTPATH"DetectUsedThemes.txt";

	///Create an instance of presentation document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load file
	ppt->LoadFromFile(inputFile.c_str());

	wofstream outFile(outputFile, ios::out);
	std::wstring themeName = L"";
	outFile << "This is the name list of the used theme below." << endl;
	//Get the theme name of each slide in the document
	for (int s = 0; s < ppt->GetSlides()->GetCount(); s++)
	{
		intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(s);
		themeName = slide->GetTheme()->GetName();
		outFile << themeName.c_str() << endl;
	}
	//Save to the text document
	outFile.close();
		
}
