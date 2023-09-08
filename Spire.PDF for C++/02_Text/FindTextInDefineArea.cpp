#include "../pch.h"

using namespace Spire::Pdf;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path+L"SampleB_1.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile= output_path+L"FindTextInDefineArea.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<RectangleF> rectg = new RectangleF(0, 0, 300, 300);
	intrusive_ptr<PdfTextFindCollection> findCollection = doc->GetPages()->GetItem(0)
		->FindText(rectg, L"Spire", Find_TextFindParameter::WholeWord);
	intrusive_ptr<PdfTextFindCollection> findCollectionOut = doc->GetPages()->GetItem(0)
		->FindText(rectg, L"PDF", Find_TextFindParameter::WholeWord);
	for(intrusive_ptr<PdfTextFind> find : findCollection->GetFinds())
	{
		find->ApplyHighLight(Color::GetGreen());
	}
	for(intrusive_ptr<PdfTextFind> findOut : findCollectionOut->GetFinds())
	{
		findOut->ApplyHighLight(Color::GetYellow());
	}
	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}


