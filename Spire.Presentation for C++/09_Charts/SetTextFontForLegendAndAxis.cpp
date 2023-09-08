#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_2.pptx";
	wstring outputFile = OUTPUTPATH"SetTextFontForLegendAndAxis.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Set the font for the text on Chart Legend area.
	chart->GetChartLegend()->GetTextProperties()->GetParagraphs()->GetItem(0)->GetDefaultCharacterProperties()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Green);
	chart->GetChartLegend()->GetTextProperties()->GetParagraphs()->GetItem(0)->GetDefaultCharacterProperties()->SetLatinFont(new TextFont(L"Arial Unicode MS"));

	//Set the font for the text on Chart Axis area.
	chart->GetPrimaryCategoryAxis()->GetTextProperties()->GetParagraphs()->GetItem(0)->GetDefaultCharacterProperties()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Red);
	chart->GetPrimaryCategoryAxis()->GetTextProperties()->GetParagraphs()->GetItem(0)->GetDefaultCharacterProperties()->GetFill()->SetFillType(FillFormatType::Solid);
	chart->GetPrimaryCategoryAxis()->GetTextProperties()->GetParagraphs()->GetItem(0)->GetDefaultCharacterProperties()->SetFontHeight(10);
	chart->GetPrimaryCategoryAxis()->GetTextProperties()->GetParagraphs()->GetItem(0)->GetDefaultCharacterProperties()->SetLatinFont(new TextFont(L"Arial Unicode MS"));

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
