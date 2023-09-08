#include "../pch.h"


using namespace Spire::Pdf;
using namespace std;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreatePolylineAnnotation.pdf";


	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->Add();
	intrusive_ptr<RectangleF> rt = new RectangleF(0, 80, 200, 200);

	wstring inputFile = input_path +L"CreatePdf3DAnnotation.u3d";
	intrusive_ptr<Pdf3DAnnotation> annotation = new Pdf3DAnnotation(rt, inputFile.c_str());
	annotation->SetActivation(new Pdf3DActivation());
	annotation->GetActivation()->SetActivationMode(Pdf3DActivationMode::PageOpen);
	intrusive_ptr<Pdf3DView> View = new Pdf3DView();
	View->SetBackground(new Pdf3DBackground(new PdfRGBColor(Spire::Common::Color::GetPurple())));
	View->SetViewNodeName(L"3DAnnotation");
	View->SetRenderMode(new Pdf3DRendermode(Pdf3DRenderStyle::Solid));
	View->SetInternalName(L"3DAnnotation");
	View->SetLightingScheme(new Pdf3DLighting());
	View->GetLightingScheme()->SetStyle(Pdf3DLightingStyle::Day);
	annotation->GetViews()->Add(View);
	page->GetAnnotationsWidget()->Add(annotation);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();
}
