#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ColumnStacked.pptx";
	wstring outputFile = OUTPUTPATH"SetPercentageForLabels.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	for (int i = 0; i < chart->GetSeries()->GetCount(); i++)
	{
		intrusive_ptr<ChartSeriesDataFormat> series = chart->GetSeries()->GetItem(i);
		//Get the total number
		float total = 0;
		for (int k = 0; k < series->GetValues()->GetCount(); k++)
		{
			total += stof(series->GetValues()->GetItem(k)->GetText());
		}
		for (int j = 0; j < series->GetValues()->GetCount(); j++)
		{
			//Get the percent
			float tp = stof(series->GetValues()->GetItem(j)->GetText()) / total;
			int dataPontPercent = tp > 0.0 ? (tp + 0.00005) * 10000 : (tp - 0.00005) * 10000;
			//Add datalabels
			intrusive_ptr<ChartDataLabel> label = series->GetDataLabels()->Add();
			label->SetLabelValueVisible(true);
			//Set the percent text for the label
			wstringstream ss;
			ss.precision(2);
			ss << fixed;
			ss << dataPontPercent / 100.0;
		std:wstring text = ss.str() + L"%";
			label->GetTextFrame()->GetParagraphs()->GetItem(0)->SetText(text.c_str());
			label->GetTextFrame()->GetParagraphs()->GetItem(0)->GetTextRanges()->GetItem(0)->SetFontHeight(12);
		}
	}
	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
