#include "../pch.h"

using namespace Spire::Pdf;
using namespace std;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"3D.pdf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"3d0.u3d";

	intrusive_ptr<PdfDocument> doc = new PdfDocument();
	doc->LoadFromFile(inputFile.c_str());
	intrusive_ptr<PdfPageBase> page = doc->GetPages()->GetItem(0);
	intrusive_ptr<PdfAnnotationCollection> annot = page->GetAnnotationsWidget();
	int count = 0;
	for (int i = 0; i < annot->GetCount(); i++) {
		wchar_t nm_w[100];
		swprintf(nm_w, 100, L"%hs", typeid(Pdf3DAnnotation).name());
		LPCWSTR_S newName = nm_w;
		if (wcscmp(newName, annot->GetItem(i)->GetInstanceTypeName()) == 0) {
			intrusive_ptr<Pdf3DAnnotation> annot3D = Object::Convert<Pdf3DAnnotation>(annot->GetItem(i));
			intrusive_ptr<Stream> stream = annot3D->Get_3DData();
			stream->Save(outputFile.c_str());
		}
	}
	doc->Close();
}

