#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ExtractParagraphBasedOnStyle.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExtractParagraphBasedOnStyle.txt";
	

	//Create a new document
	intrusive_ptr<Document> document = new Document();
	std::wstring styleName1 = L"Heading1";
	wstring* style1Text = new wstring();
	//Load file from disk
	document->LoadFromFile(inputFile.c_str());
	style1Text->append(L"The following is the content of the paragraph with the style name " + styleName1 + L": \n");
	//Extrct paragraph based on style
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(j);
			if (paragraph->GetStyleName() != nullptr && paragraph->GetStyleName() == styleName1)
			{
				style1Text->append(paragraph->GetText());
			}
		}
	}

	std::wofstream write(outputFile);
	
	write << style1Text->c_str();
	write.close();
	document->Close();
	delete style1Text;
}
