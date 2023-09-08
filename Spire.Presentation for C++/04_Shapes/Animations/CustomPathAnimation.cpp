#include "pch.h"

using namespace Spire::Presentation;

int main()
{
	wstring outputFile = OUTPUTPATH"CustomPathAnimation.pptx";

	//Create PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();

	//Add shape
	intrusive_ptr<IAutoShape> shape = ppt->GetSlides()->GetItem(0)->GetShapes()->AppendShape(ShapeType::Rectangle, new RectangleF(0, 0, 200, 200));

	//Add animation
	intrusive_ptr<AnimationEffect> effect = ppt->GetSlides()->GetItem(0)->GetTimeline()->GetMainSequence()->AddEffect(shape, AnimationEffectType::PathUser);
	intrusive_ptr<CommonBehaviorCollection> common = effect->GetCommonBehaviorCollection();
	intrusive_ptr<AnimationMotion> motion = Object::Dynamic_cast<AnimationMotion>(common->GetItem(0));
	motion->SetOrigin(AnimationMotionOrigin::Layout);
	motion->SetPathEditMode(AnimationMotionPathEditMode::Relative);

	//Add moin path
	intrusive_ptr<MotionPath> moinPath = new MotionPath();
	moinPath->Add(MotionCommandPathType::MoveTo, { new PointF(0, 0) }, MotionPathPointsType::CurveAuto, true);
	moinPath->Add(MotionCommandPathType::LineTo, { new PointF(0.1f, 0.1f) }, MotionPathPointsType::CurveAuto, true);
	moinPath->Add(MotionCommandPathType::LineTo, { new PointF(-0.1f, 0.2f) }, MotionPathPointsType::CurveAuto, true);
	moinPath->Add(MotionCommandPathType::End, {}, MotionPathPointsType::CurveStraight, true);
	motion->SetPath(moinPath);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
}

