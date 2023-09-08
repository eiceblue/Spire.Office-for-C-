
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ChartSample4.pptx";
	wstring outputFile =OUTPUTPATH"FillPictureInChartMarker.pptx";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Get the chart.
	intrusive_ptr<IChart> chart = Object::Dynamic_cast<IChart>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));


	intrusive_ptr<Stream> stream = new Stream(DATAPATH"Logo.png");
	intrusive_ptr<IImageData> IImage = ppt->GetImages()->Append(stream);

	//Create a ChartDataPoint object and specify the index
	intrusive_ptr<ChartDataPoint> dataPoint = new ChartDataPoint(chart->GetSeries()->GetItem(0));
	dataPoint->SetIndex(0);

	//Fill picture in marker
	dataPoint->GetMarkerFill()->GetFill()->SetFillType(FillFormatType::Picture);
	dataPoint->GetMarkerFill()->GetFill()->GetPictureFill()->GetPicture()->SetEmbedImage(IImage);

	//Set marker size
	dataPoint->SetMarkerSize(20);

	//Add the data point in series
	chart->GetSeries()->GetItem(0)->GetDataPoints()->Add(dataPoint);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}
