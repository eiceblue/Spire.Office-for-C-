
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"TrendlineEquation.pptx";
	wstring outputFile = OUTPUTPATH"ChangesForTrendLineEquation.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();
	//Load the file from disk.
	presentation->LoadFromFile(inputFile.c_str());

	//Get chart on the first slide
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Get the first trendline 
	intrusive_ptr<ITrendlines> trendline = chart->GetSeries()->GetItem(0)->GetTrendLines()[0];

	//Change font size for trendline Equation text
	intrusive_ptr<ParagraphCollection> ps = trendline->GetTrendLineLabel()->GetTextFrameProperties()->GetParagraphs();
	for (int i = 0; i < ps->GetCount(); i++)
	{
		intrusive_ptr<TextParagraph> para = ps->GetItem(i);
		para->GetDefaultCharacterProperties()->SetFontHeight(20);
		for (int j = 0; j < para->GetTextRanges()->GetCount(); j++)
		{
			para->GetTextRanges()->GetItem(j)->SetFontHeight(20);
		}
	}

	//Change position for trendline Equation
	trendline->GetTrendLineLabel()->SetOffsetX(-0.1f);
	trendline->GetTrendLineLabel()->SetOffsetY(-0.05f);

	//Save to file.
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
}
