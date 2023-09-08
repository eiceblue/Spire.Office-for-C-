#include "pch.h"

using namespace Spire::Common;
using namespace Spire::Pdf;
using namespace std;

int main()
{
	wstring inputFile = DATAPATH"DocumentsLinks.pdf";
	wstring outputFile = OUTPUTPATH"LaunchFileInNewWindow.pdf";

	//Load old PDF from disk.
	intrusive_ptr<PdfDocument> pdf = new PdfDocument();

	pdf->LoadFromFile(inputFile.c_str());

	//Define the variables
	std::vector<intrusive_ptr<PdfTextFind>> finds;
	std::vector<std::wstring> test = { L"Spire.PDF" };

	//Traverse the pages
	for (int j = 0; j < pdf->GetPages()->GetCount(); j++)
	{
		intrusive_ptr<PdfPageBase> page = pdf->GetPages()->GetItem(j);
		for (int i = 0; i < test.size(); i++)
		{
			//Find the defined string
			finds = page->FindText(test[i].c_str(), Find_TextFindParameter::WholeWord)->GetFinds();

			//Traverse the finds
			for (auto find : finds)
			{
				intrusive_ptr<PdfLaunchAction> launchAction = new PdfLaunchAction(outputFile.c_str(), PdfFilePathType::Relative);

				//Set open document in a new window
				launchAction->SetIsNewWindow(true);

				//Add annotation
				intrusive_ptr<RectangleF> rect = new RectangleF(find->GetPosition()->GetX(), find->GetPosition()->GetY(), find->GetSize()->GetWidth(), find->GetSize()->GetHeight());
				intrusive_ptr<PdfActionAnnotation> annotation = new PdfActionAnnotation(rect, launchAction);
				intrusive_ptr<PdfPageWidget> pageWg = Object::Convert<PdfPageWidget>(page);
				pageWg->GetAnnotationsWidget()->Add(annotation);

			}
		}
	}

	//Save the document
	pdf->SaveToFile(outputFile.c_str());
	pdf->Close();
}
