#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	
	wstring inputFile_1 = DATAPATH"CloneMaster1.pptx";
	wstring inputFile_2 = DATAPATH"CloneMaster2.pptx";
	wstring outputFile = OUTPUTPATH"ClonePPTMasterToAnother.pptx";

	//Load PPT1 from disk
	intrusive_ptr<Presentation> presentation1 = new Presentation();
	presentation1->LoadFromFile(inputFile_1.c_str());

	//Load PPT2 from disk
	intrusive_ptr<Presentation> presentation2 = new Presentation();
	presentation2->LoadFromFile(inputFile_2.c_str());

	//Add masters from PPT1 to PPT2
	for (int m = 0; m < presentation1->GetMasters()->GetCount(); m++)
	{
		intrusive_ptr<IMasterSlide> masterSlide = presentation1->GetMasters()->GetItem(m);
		presentation2->GetMasters()->AppendSlide(masterSlide);
	}

	//Save the document
	presentation2->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
		
}
