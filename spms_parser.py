# READS THE CONTENTS OF SPMS XML (Volkers combined version)

import xmltodict
import dateparser


def get_paper(paper):
    title = paper['title']
    abstract = paper['abstract']
    code = paper['program_codes']['program_code']['code']['#text']

    start_time = None
    if 'start_time' in paper['program_codes']['program_code']:
        start_time = paper['program_codes']['program_code']['start_time']

    duration = None
    if 'duration' in paper['program_codes']['program_code']:
        duration = paper['program_codes']['program_code']['duration']

    if isinstance(paper['contributors']['contributor'], list):
        all_authors = list(map(lambda contributor: get_person(contributor), paper['contributors']['contributor']))
        authors = list(filter(lambda person: person['person_type'] == 'Primary Author', all_authors))
    else:
        all_authors = [get_person(paper['contributors']['contributor'])]
        authors = all_authors

    funding = None
    if 'agency' in paper:
        funding = paper['agency']

    footnote = None
    if 'footnote' in paper:
        footnote = paper['footnote']

    return {
        'title': title,
        'abstract': abstract,
        'code': code,
        'start_time': start_time,
        'duration': duration,
        'authors': authors,
        'all_authors': all_authors,
        'funding': funding,
        'footnote': footnote
    }


def get_institute_name(institute):
    name = institute['name1']
    if 'name2' in institute:
        name += ', ' + institute['name2']

    if 'full_name' in institute:
        if isinstance(institute['full_name'], dict):
            if '@abbrev' in institute['full_name']:
                name = institute['full_name']['@abbrev']
            elif '#text' in institute['full_name']:
                name = institute['full_name']['#text']
            if '@type' in institute['full_name']:
                if institute['full_name']['@type'] == "Industry":
                    name = institute['name1']
                    name += ', ' + institute['town'] + ', ' + institute['country_code']['@abbrev']
                elif "University" in institute['full_name']['@type']:
                    if '@abbrev' in institute['full_name']:
                        name = institute['full_name']['@abbrev']
                    else:
                        name = institute['name1']
        else:
            name = institute['full_name']

    return name


def get_institute(institute):
    name = get_institute_name(institute)
    website = None
    if 'URL' in institute:
        website = institute['URL']
    location = institute['town']
    country = institute['country_code']['@abbrev']
    return {
        'name': name,
        'website': website,
        'location': location,
        'country': country
    }


def get_person(person):
    author_id = person['author_id']
    person_type = None
    if '@type' in person:
        person_type = person['@type']
    first_name = person['fname']
    last_name = person['lname']

    middle_name = None
    if 'mname' in person:
        middle_name = person['mname']

    initials = person['iname']

    if isinstance(person['institutions']['institute'], list):
        institutes = list(map(lambda institute: get_institute(institute), person['institutions']['institute']))
    else:
        institutes = [get_institute(person['institutions']['institute'])]

    return {
        'id': author_id,
        'person_type': person_type,
        'first_name': first_name,
        'middle_name': middle_name,
        'last_name': last_name,
        'initials': initials,
        'institutes': institutes
    }


def get_session(session):
    general_type = session['@type']
    code = session['name']['@abbr']
    name = session['name']['#text']

    # Times
    date = session['date']['#text']
    stime = session['date']['@btime']
    etime = session['date']['@etime']
    start = dateparser.parse(date + ' ' + stime[0:2] + ':' + stime[2:4])
    end = dateparser.parse(date + ' ' + etime[0:2] + ':' + etime[2:4])

    # Location
    location_type = session['location']['@type']
    location_name = session['location']['#text']

    chair = None
    if 'chairs' in session:
        if 'chair' in session['chairs']:
            chair = get_person(session['chairs']['chair'])

    papers = []
    if 'papers' in session and session['papers'] is not None:
        if isinstance(session['papers']['paper'], list):
            papers = list(map(lambda paper: get_paper(paper), session['papers']['paper']))
        else:
            papers = [get_paper(session['papers']['paper'])]

    return { 
        'general_type': general_type,
        'code': code,
        'name': name,
        'date': date,
        'start': start,
        'end': end,
        'chair': chair,
        'papers': papers,
        'location': {
            'type': location_type,
            'name': location_name
        }
    }


def get_papers(sessions):
    # get all papers in a list
    papers = []
    for session in sessions:
        for paper in session['papers']:
            paper['session'] = session
            papers.append(paper)

    return papers


def contains_paper(papers, code):
    for paper in papers:
        if paper['code'] == code:
            return True
    return False


def get_authors(papers):
    authors = {}
    for paper in papers:
        for author in paper['all_authors']:
            if author['id'] in authors:
                if not contains_paper(authors[author['id']]['papers'], paper['code']):
                    authors[author['id']]['papers'].append(paper)
            else:
                authors[author['id']] = author
                authors[author['id']]['papers'] = [paper]

    return authors.values()


def get_sessions(data):
    return list(map(lambda session: get_session(session), data['conference']['session']))


def get_data(filename):
    with open(filename) as xml_file:
        data = xmltodict.parse(xml_file.read())
    xml_file.close()
    return data


def get_all(filename):
    data = get_data(filename)
    sessions = get_sessions(data)
    papers = get_papers(sessions)
    authors = get_authors(papers)

    return sessions, papers, authors
