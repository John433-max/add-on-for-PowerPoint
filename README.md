# Groq OpenAI PowerPoint Add-in

This PowerPoint add-in integrates Groq AI and OpenAI to fetch and display data within your PowerPoint presentations.

## Installation

1. **Clone the Repository**:
    ```bash
    git clone https://github.com/John433-max/add-on-for-PowerPoint.git
    cd add-on-for-PowerPoint
    ```

2. **Start the Local Server**:
    ```bash
    python3 -m http.server 3000
    ```

3. **Upload the Add-in to PowerPoint**:
    - Open Microsoft PowerPoint.
    - Go to `Insert` > `My Add-ins` > `Manage My Add-ins` > `Upload My Add-in`.
    - Upload the `manifest.xml` file from the cloned repository.

## Usage

1. **Open the Add-in**:
    - The add-in should appear in the task pane.
    - Click the "Fetch Data from Groq AI and OpenAI" button to fetch and display data.

2. **Fetch Data**:
    - The add-in will fetch data from Groq AI and OpenAI and display it in the task pane.

## Deletion

1. **Remove the Add-in from PowerPoint**:
    - Go to `Insert` > `My Add-ins` > `Manage My Add-ins`.
    - Find the Groq OpenAI Add-in and remove it.

2. **Stop the Local Server**:
    - If you started the local server, stop it by pressing `Ctrl+C` in the terminal.

3. **Delete the Cloned Repository**:
    ```bash
    cd ..
    rm -rf add-on-for-PowerPoint
    ```

## License

This project is licensed under the MIT License.
