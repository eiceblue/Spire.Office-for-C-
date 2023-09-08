
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChartSample2.pptx";
	wstring outputFile = OUTPUTPATH"ChangeFontSizeForChartDataTable.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get chart on the first slide
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	chart->SetHasDataTable(true);

	//Add a new paragraph in data table
	intrusive_ptr<TextParagraph> tp = new TextParagraph();
	chart->GetChartDataTable()->GetText()->GetParagraphs()->Append(tp);
	//Change the font size
	chart->GetChartDataTable()->GetText()->GetParagraphs()->GetItem(0)
		->GetDefaultCharacterProperties()->SetFontHeight(15);
	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
