#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ShapeToImage.pptx";
	wstring outputFile = OUTPUTPATH;

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());
	intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(0);
	intrusive_ptr<ShapeCollection> shapes = slide->GetShapes();
	for (int i = 0; i < shapes->GetCount(); i++)
	{
		intrusive_ptr<Stream> image = shapes->SaveAsImage(i);
		wstring filename = outputFile + L"//ShapeToImage-" + to_wstring(i) + L".png";
		image->Save(filename.c_str());
		image->Dispose();
	}
	////delete presentation;;

}

