#include "../pch.h"

using namespace std;
using namespace Spire::Pdf;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToImage.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ConvertAllPagesToEMF\\";

	intrusive_ptr<Spire::Common::Stream> image = new Spire::Common::Stream();
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();
	pdf->LoadFromFile(inputFile.c_str());
	//Save to images
	for (int i = 0; i < pdf->GetPages()->GetCount(); i++)
	{
		wstring s = to_wstring(i);
		wstring filename = outputFile + L"ToEMF-img-" + s + L".emf";
		//var image
		image = pdf->SaveAsImage(i, Spire::Pdf::PdfImageType::Bitmap);
		image->Save(filename.c_str());
		image->Close();
	}
	pdf->Close();
}







