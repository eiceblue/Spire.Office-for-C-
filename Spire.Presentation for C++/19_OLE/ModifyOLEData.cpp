#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ModifyOLEData.pptx";
	wstring outputFile = OUTPUTPATH"ModifyOLEData.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Load document from disk
	ppt->LoadFromFile(inputFile.c_str());

	//Loop through the slides and shapes
	intrusive_ptr<SlideCollection> slides = ppt->GetSlides();
	for (int i = 0; i < slides->GetCount(); i++)
	{
		intrusive_ptr<ISlide> slide = slides->GetItem(i);
		for (int j = 0; j < slide->GetShapes()->GetCount(); j++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(j);
			if (Object::Dynamic_cast<Spire::Presentation::IOleObject>(shape) != nullptr)
			{
				//Find OLE object
				intrusive_ptr<Spire::Presentation::IOleObject> oleObject = Object::Dynamic_cast<Spire::Presentation::IOleObject>(shape);

				intrusive_ptr<Stream> stream = oleObject->GetDataStream();
				//Get its data and write to file
				if (wcscmp(oleObject->GetProgId(), L"PowerPoint.Show.12") == 0)
				{
					//Load the PPT stream
					intrusive_ptr<Presentation> ppt = new Presentation();
					ppt->LoadFromStream(stream, FileFormat::Auto);
					//Append an image in slide
					wstring inputFile1 = DATAPATH"Logo.png";
					ppt->GetSlides()->GetItem(0)->GetShapes()->AppendEmbedImage(ShapeType::Rectangle, inputFile1.c_str(), new RectangleF(50, 50, 100, 100));
					intrusive_ptr<Stream> stream2 = new Stream();
					ppt->SaveToFile(stream2, FileFormat::Pptx2013);
					stream2->SetPosition(0);
					//Modify the data
					oleObject->SetDataStream(stream2);
				}
			}
		}
	}

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}

