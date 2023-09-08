#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	wstring inputFile = DATAPATH"AddSection.pptx";
	wstring outputFile = OUTPUTPATH"GetSectionIndex.txt";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Section> section = ppt->GetSectionList()->GetItem(0);

	int index = ppt->GetSectionList()->IndexOf(section);

	//Save to file.
	std::wofstream out;
	out.open(outputFile);
	out.flush();
	out << L"index:" + std::to_wstring(index);
	out.close();

}
