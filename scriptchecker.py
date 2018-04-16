
import adminfile


# Function that check the times are all different
def check_times(detailed_log_path):
    f = open(detailed_log_path, 'r')
    line = f.readline()
    keysfile_adds = ""
    check_list = []
    while line != '':
        # Get attributes from the log line.
        date, time, program_name, window_id, username, window_title, logged_keys = line.split("|")
        keysfile_adds += date + "|" + time + "|" + program_name + "|" + window_id + "|" + username + "|" + window_title + "|"

        logged_keys = logged_keys.split(" ")

        ms, key, msg, x, y = logged_keys[0].split(",")
        if check_list.__len__() != 0 and check_list.__contains__(ms):
            keysfile_adds += str(int(ms) + 1) + "," + key + "," + msg + "," + x + "," + y
            check_list.append(str(int(ms) + 1))
        else:
            keysfile_adds += logged_keys[0]
            check_list.append(ms)

        for i in range(1, logged_keys.__len__()):  # It is taken as tautology that always starts with a down
            logged_key = logged_keys[i]
            # print logged_key
            if logged_key == "\n":
                keysfile_adds += "\n"
            else:
                ms, key, msg, x, y = logged_key.split(",")
                if check_list.__contains__(ms):
                    keysfile_adds += " " + str(int(ms) + 1) + "," + key + "," + msg + "," + x + "," + y
                    check_list.append(str(int(ms) + 1))
                else:
                    keysfile_adds += " " + ms + "," + key + "," + msg + "," + x + "," + y
                    check_list.append(ms)

        line = f.readline()

    adminfile.writefile(detailed_log_path, keysfile_adds)


def arranger(detailed_log_path, click_images_log_path):

    keysfile = adminfile.readfile(detailed_log_path)
    keysfile = keysfile.replace("key down", "key_down,-1,-1")
    keysfile = keysfile.replace("key up", "key_up,-1,-1")
    keysfile = keysfile.replace("key sys down", "key_down,-1,-1")
    keysfile = keysfile.replace("key sys up", "key_up,-1,-1")
    keysfile = keysfile.replace("Lcontrol", "ctrll")
    keysfile = keysfile.replace("Rcontrol", "ctrlr")
    keysfile = keysfile.replace("Left", "leftarrow")
    keysfile = keysfile.replace("Up", "uparrow")
    keysfile = keysfile.replace("Right", "rightarrow")
    keysfile = keysfile.replace("Down", "downarrow")
    keysfile = keysfile.replace("Back", "backspace")
    keysfile = keysfile.replace("Delete", "supr")
    keysfile = keysfile.replace("Capital", "mayus")
    keysfile = keysfile.replace("Oem_Comma", "comma")
    keysfile = keysfile.replace("Oem_Period", "period")
    keysfile = keysfile.replace("Oem_Plus", "plus")

    pipe = 0
    comma = 0
    keysfile_adds = ""
    for i in range(0, keysfile.__len__()):
        char = keysfile[i]
        if char == "\n":
                pipe = 0
                keysfile_adds += char
        elif pipe == 6:
            if char == " ":
                keysfile_adds += char
                comma = 0
            elif char == ",":
                keysfile_adds += char
                comma += 1
            elif comma == 1:
                keysfile_adds += char.lower()
            else:
                keysfile_adds += char
        elif char == "|":
            pipe += 1
            keysfile_adds += char
        else:
            keysfile_adds += char

    detailed_log_path = detailed_log_path.replace(".txt", "_new.txt")
    adminfile.writefile(detailed_log_path, keysfile_adds)

    check_times(detailed_log_path)

    clicksfile = adminfile.readfile(click_images_log_path)
    clicksfile = clicksfile.replace("mouse left up", "left_up")
    clicksfile = clicksfile.replace("mouse left down", "left_down")
    clicksfile = clicksfile.replace("mouse right up", "left_up")
    clicksfile = clicksfile.replace("mouse right down", "left_down")
    clicksfile = clicksfile.replace("mouse middle down", "middle_down")
    clicksfile = clicksfile.replace("mouse middle up", "middle_up")

    pipe = 0
    comma = 0
    clicksfile_adds = ""
    for i in range(0, clicksfile.__len__()):
        char = clicksfile[i]

        if pipe == 5 and clicksfile[i-1] == "|":
            if char == "|":
                pipe += 1
                clicksfile_adds += "|"
            else:
                clicksfile_adds += char
        elif pipe == 7:
            if comma == 4:
                if char == "g":
                    temp = clicksfile[i-3] + clicksfile[i-2] + clicksfile[i-1] + char
                    if temp == ".png":
                        comma = 0
                        clicksfile_adds += char
                    else:
                        clicksfile_adds += char
                elif char == " ":
                    clicksfile_adds += "_"
                elif char == "\n":
                    clicksfile_adds += "\\n"
                else:
                    clicksfile_adds += char
            elif char == ",":
                clicksfile_adds += char
                comma += 1
            else:
                clicksfile_adds += char
        elif char == "|":
            pipe += 1
            clicksfile_adds += char
        elif char == "\n":
            pipe = 0
            clicksfile_adds += char
        else:
            clicksfile_adds += char

    click_images_log_path = click_images_log_path.replace(".txt", "_new.txt")
    adminfile.writefile(click_images_log_path, clicksfile_adds)
