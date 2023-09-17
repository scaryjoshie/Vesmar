from subprocess import check_call

# On windows use:
def copy2clip(txt):
    cmd='echo '+txt.strip()+'|clip'
    return check_call(cmd, shell=True)

while True:
    text = input("Paste Text: ")
    text = text.replace(",", "")

    split_text = text.split()

    merged = "'" + "	'".join(split_text)

    copy2clip(merged)

    print("Copied!")