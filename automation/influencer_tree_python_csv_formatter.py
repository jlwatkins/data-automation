import requests
response = requests.get("https://www.socio-ip-domain.com/connections_tree?event_id=309")
data = response.json()
with open("./influencer_tree_simple_output.csv", 'wb') as f:
    line = str()
    for basic_item in data:
        this_person = basic_item[0]
        his_friends_list = basic_item[1]
        line += str(this_person) + "," + his_friends_list[0] + "\n"
        for person in his_friends_list[1:]:
            line += "," + person + "\n"

    f.write(line.encode("UTF-8"))
