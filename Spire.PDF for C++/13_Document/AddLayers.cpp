#include "pch.h"

using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main()
{
	wstring inputFile = DATAPATH"AddLayers.pdf";
	wstring outputFile = OUTPUTPATH"AddLayers.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());

	//LPCWSTR_S aaa = new PdfJavaScript->GetNumberFormatString();

	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);

	//create a layer named "red line"
	intrusive_ptr<PdfLayer> layer = doc->GetLayers()->AddLayer(L"red line", PdfVisibility::On);
	intrusive_ptr<PdfCanvas> pcA = layer->CreateGraphics(page->GetCanvas());
	//PdfPen tempVar(PdfBrushes::GetRed(), 2);
	intrusive_ptr<PdfPen> tempVar = new PdfPen(PdfBrushes::GetRed(), 2);
	pcA->DrawLine(tempVar, new PointF(100, 350), new PointF(300, 350));

	//create a layer named "blue line"
	layer = doc->GetLayers()->AddLayer(L"blue line");
	intrusive_ptr<PdfCanvas> pcB = layer->CreateGraphics(doc->GetPages()->GetItem(0)->GetCanvas());
	//PdfPen tempVar2(PdfBrushes::Blue, 2);
	intrusive_ptr<PdfPen> tempVar2 = new PdfPen(PdfBrushes::GetBlue(), 2);
	pcB->DrawLine(tempVar2, new PointF(100, 400), new PointF(300, 400));

	//create a layer named "green line"
	layer = doc->GetLayers()->AddLayer(L"green line");
	intrusive_ptr<PdfCanvas> pcC = layer->CreateGraphics(doc->GetPages()->GetItem(0)->GetCanvas());
	//PdfPen tempVar3(PdfBrushes::Green, 2);
	intrusive_ptr<PdfPen> tempVar3 = new PdfPen(PdfBrushes::GetGreen(), 2);
	pcC->DrawLine(tempVar3, new PointF(100, 450), new PointF(300, 450));
	//Save the document
	doc->SaveToFile(outputFile.c_str());
	doc->Close();
}
