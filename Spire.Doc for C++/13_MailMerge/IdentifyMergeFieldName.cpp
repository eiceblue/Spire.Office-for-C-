#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"IdentifyMergeFieldNames.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"IdentifyMergeFieldName.txt";
	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the collection of group names.
	std::vector<LPCWSTR_S> GroupNames = document->GetMailMerge()->GetMergeGroupNames();

	//Get the collection of merge field names in a specific group.
	std::vector<LPCWSTR_S> MergeFieldNamesWithinRegion = document->GetMailMerge()->GetMergeFieldNames(L"Products");

	//Get the collection of all the merge field names.
	std::vector<LPCWSTR_S> MergeFieldNames = document->GetMailMerge()->GetMergeFieldNames();

	wstring* content = new wstring();
	content->append(L"----------------Group Names-----------------------------------------");
	content->append(L"\n");
	for (int i = 0; i < GroupNames.size(); i++)
	{
		content->append(GroupNames[i]);
		content->append(L"\n");
	}

	content->append(L"----------------Merge field names within a specific group-----------");
	content->append(L"\n");
	for (int j = 0; j < MergeFieldNamesWithinRegion.size(); j++)
	{
		content->append(MergeFieldNamesWithinRegion[j]);
		content->append(L"\n");
	}

	content->append(L"----------------All of the merge field names------------------------");
	content->append(L"\n");
	for (int k = 0; k < MergeFieldNames.size(); k++)
	{
		content->append(MergeFieldNames[k]);
		content->append(L"\n");
	}

	//Save to file.
	std::wofstream foo(outputFile);
	foo << content->c_str();
	foo.close();
	
	document->Close();
	
	delete content;
}
