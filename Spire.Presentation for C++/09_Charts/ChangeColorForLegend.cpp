
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChartSample2.pptx";
	wstring outputFile = OUTPUTPATH"ChangeColorForLegend.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get chart on the first slide
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Change the fill color
	chart->GetChartLegend()->GetTextProperties()->GetParagraphs()->GetItem(0)
		->GetDefaultCharacterProperties()->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetChartLegend()->GetTextProperties()->GetParagraphs()->GetItem(0)
		->GetDefaultCharacterProperties()->GetFill()->GetSolidColor()->SetColor(Color::GetBlue());
	//Use italic for the paragraph
	chart->GetChartLegend()->GetTextProperties()->GetParagraphs()->GetItem(0)
		->GetDefaultCharacterProperties()->SetIsItalic(TriState::True);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
