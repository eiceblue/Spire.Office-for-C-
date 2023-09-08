#include "../pch.h"

using namespace std;
using namespace Spire::Pdf;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddSVGToPDF.pdf";
	wstring fn1 = DATAPATH;

	//Create a new PDF document.
	intrusive_ptr<PdfDocument> existingPDF = new PdfDocument();

	existingPDF->LoadFromFile((fn1 + L"SampleB_1.pdf").c_str());
	//Create a new PDF document.
	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	//Load the SVG file
	doc->LoadFromSvg((fn1 + L"template.svg").c_str());
	//Create template
	intrusive_ptr<PdfTemplate> pdftemplate = doc->GetPages()->GetItem(0)->CreateTemplate();

	float x = 50;
	float y = 250;

	intrusive_ptr<PointF> pf = new PointF(x, y);

	existingPDF->GetPages()->GetItem(0)->GetCanvas()->DrawTemplate(doc->GetPages()->GetItem(0)->CreateTemplate(),
		pf, new Spire::Common::SizeF(200, 300));

	existingPDF->SaveToFile((outputFile).c_str(), FileFormat::PDF);
	existingPDF->Close();
	doc->Close();
}




