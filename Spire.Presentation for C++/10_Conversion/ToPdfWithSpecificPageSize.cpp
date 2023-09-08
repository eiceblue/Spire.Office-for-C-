#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ToPDF.pptx";
	wstring outputFile = OUTPUTPATH"ToPdfWithSpecificPageSize.pdf";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Set A4 page size
	ppt->GetSlideSize()->SetType(SlideSizeType::A4);

	//Set landscape orientation
	ppt->GetSlideSize()->SetOrientation(SlideOrienation::Landscape);

	//Save the document to HTML format
	ppt->SaveToFile(outputFile.c_str(), FileFormat::PDF);

}
