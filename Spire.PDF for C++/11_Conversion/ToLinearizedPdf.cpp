#include "pch.h"

using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"PDFTemplate_N.pdf";
	wstring outputFile = OUTPUTPATH"ToLinearizedPdf.pdf";

	intrusive_ptr<PdfToLinearizedPdfConverter> converter = new PdfToLinearizedPdfConverter(inputFile.c_str());
	converter->ToLinearizedPdf(outputFile.c_str());
}
