#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"BlankSample.pptx";
	wstring outputFile = OUTPUTPATH"AddSection.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the second slide
	intrusive_ptr<ISlide> slide = ppt->GetSlides()->GetItem(1);

	//Append section with section name at the end
	ppt->GetSectionList()->Append(L"E-iceblue01");
	//Add section with slide
	ppt->GetSectionList()->Add(L"section1", slide);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}
