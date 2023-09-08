#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"ExtractOLEObject.pptx";
	wstring outputFile_px = OUTPUTPATH"ExtractOLEObject.pptx";
	wstring outputFile_p = OUTPUTPATH"ExtractOLEObject.ppt";
	wstring outputFile_xls = OUTPUTPATH"tractOLEObject.xls";
	wstring outputFile_xlsx = OUTPUTPATH"ExtractOLEObject.xlsx";
	wstring outputFile_doc = OUTPUTPATH"ExtractOLEObject.doc";
	wstring outputFile_docx = OUTPUTPATH"ExtractOLEObject.docx";

	//Create a PPT document
	intrusive_ptr<Presentation> presentation = new Presentation();

	//Load document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Loop through the slides and shapes
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		intrusive_ptr<ISlide> slide = presentation->GetSlides()->GetItem(i);
		for (int k = 0; k < slide->GetShapes()->GetCount(); k++)
		{
			intrusive_ptr<IShape> shape = slide->GetShapes()->GetItem(k);
			if (Object::Dynamic_cast<Spire::Presentation::IOleObject>(shape) != nullptr)
			{
				//Find OLE object
				intrusive_ptr<Spire::Presentation::IOleObject> oleObject = Object::Dynamic_cast<Spire::Presentation::IOleObject>(shape);

				//Get its data and write to file
				intrusive_ptr<Stream> stream = oleObject->GetDataStream();

				if (wcscmp(oleObject->GetProgId(), L"Excel.Sheet.8") == 0)
				{
					stream->Save(outputFile_xls.c_str());
				}
				//ORIGINAL LINE: case "Excel.Sheet.12":
				else if (wcscmp(oleObject->GetProgId(), L"Excel.Sheet.12") == 0)
				{
					stream->Save(outputFile_xlsx.c_str());
				}
				//ORIGINAL LINE: case "Word.Document.8":
				else if (wcscmp(oleObject->GetProgId(), L"Word.Document.8") == 0)
				{
					stream->Save(outputFile_doc.c_str());
				}
				//ORIGINAL LINE: case "Word.Document.12":
				else if (wcscmp(oleObject->GetProgId(), L"Word.Document.12") == 0)
				{
					stream->Save(outputFile_docx.c_str());
				}
				//ORIGINAL LINE: case "PowerPoint.Show.8":
				else if (wcscmp(oleObject->GetProgId(), L"PowerPoint.Show.8") == 0)
				{
					stream->Save(outputFile_p.c_str());
				}
				//ORIGINAL LINE: case "PowerPoint.Show.12":
				else if (wcscmp(oleObject->GetProgId(), L"PowerPoint.Show.12") == 0)
				{
					stream->Save(outputFile_px.c_str());
				}
				stream->Close();
			}
		}
	}
}

