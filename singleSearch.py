from linkedin_api import Linkedin
import argparse, json

parser = argparse.ArgumentParser(description='blah')
parser.add_argument('keyw', metavar='N', type=str,
                    help='keywords for linkedin search')
parser.add_argument('usn', type=str,
                    help='user')
parser.add_argument('pwd', type=str,
                    help='pwd')


args = parser.parse_args()
keyw=args.keyw
api = Linkedin(args.usn, args.pwd, refresh_cookies=True)

search = api.search_people(keywords=keyw)

print(search)

with open("searchResult.json", "w") as fh:
    json.dump(search,fh)
