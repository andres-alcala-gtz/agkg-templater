import re
import docx
import pandas
import pathlib
import datetime


def search_part_style(path: pathlib.Path) -> dict[str, str]:

    part_paragraph = {"section": "Plan Appraisal - Sample Scope: Timeline and Scheduling", "title": "Appraisal Timeline", "subtitle": "Phase 2: Conduct Appraisal", "label": "Appraisal Type", "text": "Benchmark"}

    style_paragraphs = {}
    document = docx.Document(str(path))
    for paragraph in document.paragraphs:
        if paragraph.style is not None and paragraph.text != "":
            if paragraph.style.name not in style_paragraphs:
                style_paragraphs[paragraph.style.name] = [paragraph.text]
            else:
                style_paragraphs[paragraph.style.name].append(paragraph.text)

    part_style = {}
    for part, paragraph in part_paragraph.items():
        for style, paragraphs in style_paragraphs.items():
            if paragraph in paragraphs:
                part_style[part] = style

    return part_style


def search_data_frame(path: pathlib.Path, part_style: dict[str, str]) -> pandas.DataFrame:

    def _push(data_frame: pandas.DataFrame, dictionary: dict[str, str]) -> None:
        data_frame.loc[len(data_frame)] = list(dictionary.values())

    dictionary = {part_style["section"]: "", part_style["title"]: "", part_style["subtitle"]: "", part_style["label"]: "", part_style["text"]: ""}
    data_frame = pandas.DataFrame(columns=list(dictionary.keys()))

    dictionary.update({part_style["section"]: "Plan Appraisal - Cover", part_style["label"]: "AID:", part_style["text"]: re.search("[\d]+", path.name).group()})
    _push(data_frame, dictionary)

    document = docx.Document(str(path))
    for paragraph in document.paragraphs:
        if paragraph.style is not None and paragraph.text != "":
            if paragraph.style.name in dictionary:
                dictionary.update({paragraph.style.name: paragraph.text})
            if paragraph.style.name == part_style["text"]:
                _push(data_frame, dictionary)

    return data_frame


def search_data(data_frame: pandas.DataFrame, part_style: dict[str, str]) -> dict[str, str]:

    def _loc(personnel: bool = False, index: int = 0, section: str = None, title: str = None, subtitle: str = None, label: str = None, text: str = None) -> str:
        filtered = data_frame
        if section: filtered = filtered[filtered[part_style["section"]] == section]
        if title: filtered = filtered[filtered[part_style["title"]] == title]
        if subtitle: filtered = filtered[filtered[part_style["subtitle"]] == subtitle]
        if label: filtered = filtered[filtered[part_style["label"]] == label]
        if text: filtered = filtered[filtered[part_style["text"]] == text]
        try: return filtered[part_style["subtitle" if personnel else "text"]].values[index]
        except: return "-"

    data = {}

    data["«APPRAISAL_IDENTIFIER»"] = _loc(section="Plan Appraisal - Cover", label="AID:")
    data["«APPRAISAL_TARGET»"] = _loc(section="Plan Appraisal - Cover", label="Target:")

    data["«APPRAISAL_TIMEZONE»"] = _loc(section="Plan Appraisal – Sample Scope: Appraisal Setup", label="Time Zone")
    data["«APPRAISAL_PARTNER»"] = _loc(section="Plan Appraisal – Sample Scope: Appraisal Setup", label="Partner")
    data["«APPRAISAL_OBJECTIVES»"] = _loc(section="Plan Appraisal – Sample Scope: Appraisal Setup", label="Business and Appraisal Objectives")
    data["«APPRAISAL_VIRTUAL»"] = _loc(section="Plan Appraisal – Sample Scope: Appraisal Setup", label="Use of a virtual collection technique for appraisal")

    data["«ORGANIZATION_NAME»"] = _loc(section="Plan Appraisal – Sample Scope: Organization", title="Organization", label="Name")
    data["«ORGANIZATION_NATIVE_NAME»"] = _loc(section="Plan Appraisal – Sample Scope: Organization", title="Organization", label="Native Language Name")
    data["«ORGANIZATION_CITY»"] = _loc(section="Plan Appraisal – Sample Scope: Organization", title="Organization", label="City")
    data["«ORGANIZATION_REGION»"] = _loc(section="Plan Appraisal – Sample Scope: Organization", title="Organization", label="State/Province/Region")

    data["«ORGANIZATIONAL_UNIT_NAME»"] = _loc(section="Plan Appraisal – Sample Scope: Organization", title="Organizational Unit (OU)", label="Name")
    data["«ORGANIZATIONAL_UNIT_NATIVE_NAME»"] = _loc(section="Plan Appraisal – Sample Scope: Organization", title="Organizational Unit (OU)", label="Native Language Name")

    data["«SPONSOR_NAME»"] = _loc(personnel=True, section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", text="Appraisal Sponsor")
    data["«SPONSOR_ORGANIZATION»"] = _loc(section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", subtitle=data["«SPONSOR_NAME»"], label="Organization")

    data["«INTERPRETER_NAME»"] = _loc(personnel=True, section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", text="Registered Interpreter")
    data["«INTERPRETER_ORGANIZATION»"] = _loc(section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", subtitle=data["«INTERPRETER_NAME»"], label="Organization")

    data["«COORDINATOR_NAME»"] = _loc(personnel=True, section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", text="OU Coordinator")
    data["«COORDINATOR_ORGANIZATION»"] = _loc(section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", subtitle=data["«COORDINATOR_NAME»"], label="Organization")

    data["«MEMBER0_NAME»"] = _loc(personnel=True, index=0, section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", text="Appraisal Team Member")
    data["«MEMBER0_ORGANIZATION»"] = _loc(section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", subtitle=data["«MEMBER0_NAME»"], label="Organization")

    data["«MEMBER1_NAME»"] = _loc(personnel=True, index=1, section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", text="Appraisal Team Member")
    data["«MEMBER1_ORGANIZATION»"] = _loc(section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", subtitle=data["«MEMBER1_NAME»"], label="Organization")

    data["«MEMBER2_NAME»"] = _loc(personnel=True, index=2, section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", text="Appraisal Team Member")
    data["«MEMBER2_ORGANIZATION»"] = _loc(section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", subtitle=data["«MEMBER2_NAME»"], label="Organization")

    data["«MEMBER3_NAME»"] = _loc(personnel=True, index=3, section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", text="Appraisal Team Member")
    data["«MEMBER3_ORGANIZATION»"] = _loc(section="Plan Appraisal - Sample Scope: Appraisal Personnel", title="Create Appraisal Team", subtitle=data["«MEMBER3_NAME»"], label="Organization")

    data["«PLAN_START»"] = datetime.datetime.strptime(_loc(section="Plan Appraisal - Sample Scope: Timeline and Scheduling", label="Plan Appraisal Start Date"), "%Y/%m/%d").strftime("%d-%b-%Y")
    data["«PLAN_END»"] = datetime.datetime.strptime(_loc(section="Plan Appraisal - Sample Scope: Timeline and Scheduling", label="Plan Appraisal End Date"), "%Y/%m/%d").strftime("%d-%b-%Y")
    data["«CONDUCT_START»"] = datetime.datetime.strptime(_loc(section="Plan Appraisal - Sample Scope: Timeline and Scheduling", label="Conduct Appraisal Start Date"), "%Y/%m/%d").strftime("%d-%b-%Y")
    data["«CONDUCT_END»"] = datetime.datetime.strptime(_loc(section="Plan Appraisal - Sample Scope: Timeline and Scheduling", label="Conduct Appraisal End Date"), "%Y/%m/%d").strftime("%d-%b-%Y")

    return data


def search(directory_src: pathlib.Path) -> dict[str, str]:

    path = [path for path in directory_src.glob("*") if path.suffix == ".docx" and not path.name.startswith(".")][0]

    part_style = search_part_style(path)
    data_frame = search_data_frame(path, part_style)
    data = search_data(data_frame, part_style)

    return data
