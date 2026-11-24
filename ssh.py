from netmiko import ConnectHandler, ReadTimeout


def create_connection(device):
    connection = ConnectHandler(**device)
    if device["device_type"] != "linux":
        try:
            connection.enable()
        except Exception:
            connection.disconnect()
            raise
    return connection


def get_prompt(connection):
    return connection.find_prompt()


def close(connection):
    connection.disconnect()


def run(connection, cmd):
    try:
        output = connection.send_command(cmd)
    except ReadTimeout as e:
        raise RuntimeError("Timed out on: " + cmd) from e
    lines = output.splitlines()
    for i in range(len(lines)):
        if lines[i].startswith("% "):
            raise RuntimeError("Switch error on: " + cmd + " -> " + lines[i])
    return lines


if __name__ == "__main__":
    device = {
        "device_type": "linux",
        "host": "192.168.11.213",
        "username": "admin",
        "password": "password",
    }
    connection = create_connection(device)
    try:
        print("Connection established successfully.")
        output = run(connection, "show version")
        print("\n".join(output))
    finally:
        connection.disconnect()
