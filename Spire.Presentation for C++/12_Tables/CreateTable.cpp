#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"CreateTable.pptx";
	wstring outputFile = OUTPUTPATH"CreateTable.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	std::vector<double> widths = { 100, 100, 150, 100, 100 };
	std::vector<double> heights = { 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15 };

	//Add new table to PPT
	intrusive_ptr<ITable> table = presentation->GetSlides()->GetItem(0)->GetShapes()->AppendTable(presentation->GetSlideSize()->GetSize()->GetWidth() / 2 - 275, 90, widths, heights);

	std::vector<std::vector<std::wstring>> dataStr =
	{
		{L"Name", L"Capital", L"Continent", L"Area", L"Population"},
		{L"Venezuela", L"Caracas", L"South America", L"912047", L"19700000"},
		{L"Bolivia", L"La Paz", L"South America", L"1098575", L"7300000"},
		{L"Brazil", L"Brasilia", L"South America", L"8511196", L"150400000"},
		{L"Canada", L"Ottawa", L"North America", L"9976147", L"26500000"},
		{L"Chile", L"Santiago", L"South America", L"756943", L"13200000"},
		{L"Colombia", L"Bagota", L"South America", L"1138907", L"33000000"},
		{L"Cuba", L"Havana", L"North America", L"114524", L"10600000"},
		{L"Ecuador", L"Quito", L"South America", L"455502", L"10600000"},
		{L"Paraguay", L"Asuncion", L"South America", L"406576", L"4660000"},
		{L"Peru", L"Lima", L"South America", L"1285215", L"21600000"},
		{L"Jamaica", L"Kingston", L"North America", L"11424", L"2500000"},
		{L"Mexico", L"Mexico City", L"North America", L"1967180", L"88600000"}
	};

	//Add data to table
	for (int i = 0; i < 13; i++)
	{
		for (int j = 0; j < 5; j++)
		{
			//Fill the table with data
			table->GetItem(j, i)->GetTextFrame()->SetText((dataStr[i][j]).c_str());

			//Set the Font
			table->GetItem(j, i)->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetLatinFont(new TextFont(L"Arial Narrow"));
		}
	}

	//Set the alignment of the first row to Center
	for (int i = 0; i < 5; i++)
	{
		table->GetItem(i, 0)->GetTextFrame()->GetParagraphs()->GetItem(0)->SetAlignment(TextAlignmentType::Center);
	}

	//Set the style of table
	table->SetStylePreset(TableStylePreset::LightStyle3Accent1);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);

}