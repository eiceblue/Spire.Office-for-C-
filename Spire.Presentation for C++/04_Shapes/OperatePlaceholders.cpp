#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"OperatePlaceholders.pptx";
	wstring outputFile = OUTPUTPATH"OperatePlaceholders.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Operate placeholders
	for (int j = 0; j < presentation->GetSlides()->GetCount(); j++)
	{
		intrusive_ptr<ISlide> slide = Object::Dynamic_cast<ISlide>(presentation->GetSlides()->GetItem(j));

		for (int i = 0; i < slide->GetShapes()->GetCount(); i++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(i);
			switch (shape->GetPlaceholder()->GetType())
			{
			case PlaceholderType::Media:
				shape->InsertVideo(DATAPATH"Video.mp4");
				break;

			case PlaceholderType::Picture:
				shape->InsertPicture(DATAPATH"E-iceblueLogo.png");
				break;

			case PlaceholderType::Chart:
				shape->InsertChart(ChartType::ColumnClustered);
				break;

			case PlaceholderType::Table:
				shape->InsertTable(3, 2);
				break;

			case PlaceholderType::Diagram:
				shape->InsertSmartArt(SmartArtLayoutType::BasicBlockList);
				break;
			}
		}
	}
	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

}

