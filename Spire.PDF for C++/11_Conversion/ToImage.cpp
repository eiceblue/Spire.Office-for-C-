#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring outputFile = OUTPUTPATH"ToImage/";
	wstring inputFile = DATAPATH"ToImage.pdf";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	for (int i = 0; i < doc->GetPages()->GetCount(); i++) {
		intrusive_ptr<Stream> image = doc->SaveAsImage(i);
		wstring fileName = outputFile + to_wstring(i) + L".bmp";
		image->Save(fileName.c_str());
	}
	doc->Close();
}