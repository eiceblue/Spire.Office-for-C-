#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Indent.pptx";
	std::wstring outputFile = OutputPath"Image/CustomSizeToTIFF/";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the first slide
	ISlide* slide = ppt->GetSlides()->GetItem(0);

	//Create a new PPT document
	Presentation* newPresentation = new Presentation();

	//Remove the default slide 
	newPresentation->GetSlides()->RemoveAt(0);

	//Define a new size
	SizeF* size = new SizeF(200.0, 200.0);

	//Set PPT slide size
	newPresentation->GetSlideSize()->SetSize(size);

	//Insert the slide of original PPT
	newPresentation->GetSlides()->Insert(0, slide);

	//String for output file 
	std::wstring result = outputFile + L"Output1.tiff";

	//Save to file.
	//newPresentation->SaveToFile(result.c_str(), FileFormat::Tiff);    //²»Ö§³ÖTIFF
	delete ppt;
	delete newPresentation;
}
