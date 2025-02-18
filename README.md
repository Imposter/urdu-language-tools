# Urdu Language Tools for Microsoft Word

This project is a VSTO (Visual Studio Tools for Office) plugin for Microsoft Word that provides tools for formatting Urdu poetry, including Ghazals, Nazams, and Nasars.

## Features

- Format Ghazals with specific styles and lines per verse.
- Add formatted Ghazals to the table of contents.
- Remove multiple spaces from selected text.
- Paste and format Ghazal text from the clipboard.

![Controls](docs/images/001_controls.png)
![Pasting](docs/images/002_pasting_ghazal.gif)

## Building the Project

To build the project, follow these steps:

1. Clone the repository:
    ```sh
    git clone https://github.com/imposter/UrduLanguageTools.git
    cd UrduLanguageTools
    ```

2. Open the solution file `UrduLanguageTools.sln` in Visual Studio.

3. Restore the NuGet packages:
    ```sh
    dotnet restore
    ```

4. Build the solution:
    ```sh
    dotnet build
    ```

5. Run the project to start debugging the VSTO plugin in Microsoft Word.


## Publishing the Plugin

To publish the plugin, follow these steps:

1. Create a code signing certificate using the script `ops/cert_gen.ps1`:
    ```pwsh
    .\ops\cert_gen.ps1
    ```
    This script will create a self-signed certificate and export it to a PFX file.
2. Open the solution file `UrduLanguageTools.sln` in Visual Studio.
3. Right-click on the project `UrduLanguageTools` and select `Properties`.
4. Go to the `Signing` tab and check the box `Sign the ClickOnce manifests`.
5. Click on the `Select from file...` button and select the PFX file created in step 1. The password for the PFX file is `password`.
6. Check the box `Sign the assembly`.
7. Publish the project by right-clicking on the project and selecting `Publish`.

The plugin will be published to the specified location. The published files can be copied to a shared location along with the certificate for installation on other machines.

## Installing the Plugin

To install the plugin, follow these steps:

1. Copy the published files to a shared location.
2. Install the certificate by double-clicking on the PFX file and following the installation wizard. The password for the PFX file is `password`.
3. Run the `setup.exe` file to install the plugin.

## Contributing

We welcome contributions to improve the Urdu Language Tools plugin. To contribute, follow these steps:

1. Fork the repository.

2. Create a new branch for your feature or bugfix:
    ```sh
    git checkout -b feature-or-bugfix-name
    ```

3. Make your changes and commit them with a descriptive message:
    ```sh
    git commit -am "Add new feature or fix bug"
    ```

4. Push your changes to your forked repository:
    ```sh
    git push origin feature-or-bugfix-name
    ```

5. Create a pull request to the main repository.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.