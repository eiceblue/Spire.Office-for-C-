#include "pch.h"

using namespace Spire::Common;
using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"SampleB_1.pdf";
	wstring outputFile = OUTPUTPATH"ToHTMLStream.html";

	//Create a pdf document
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	//Load an existing pdf from disk
	pdf->LoadFromFile(inputFile.c_str());
	intrusive_ptr<Stream> ms = new Stream();
	pdf->SaveToStream(ms, FileFormat::HTML);
	ms->Save(outputFile.c_str());
}
