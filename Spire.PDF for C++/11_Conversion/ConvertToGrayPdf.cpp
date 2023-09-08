#include "pch.h"


using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"ConvertToGrayPdf.pdf";
	wstring outputFile = OUTPUTPATH"ConvertToGrayPdf.pdf";

	intrusive_ptr<PdfGrayConverter> converter = new PdfGrayConverter(inputFile.c_str());
	converter->ToGrayPdf(outputFile.c_str());
}
