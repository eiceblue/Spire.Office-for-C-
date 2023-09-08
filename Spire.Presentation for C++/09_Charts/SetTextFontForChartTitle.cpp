#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"Template_Ppt_3.pptx";
	wstring outputFile = OUTPUTPATH"SetTextFontForChartTitle.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Set the font for the text on chart title area.
	chart->GetChartTitle()->GetTextProperties()->GetParagraphs()->GetItem(0)->GetDefaultCharacterProperties()->SetLatinFont(new TextFont(L"Arial Unicode MS"));
	chart->GetChartTitle()->GetTextProperties()->GetParagraphs()->GetItem(0)->GetDefaultCharacterProperties()->GetFill()->GetSolidColor()->SetKnownColor(KnownColors::Blue);
	chart->GetChartTitle()->GetTextProperties()->GetParagraphs()->GetItem(0)->GetDefaultCharacterProperties()->SetFontHeight(50);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
