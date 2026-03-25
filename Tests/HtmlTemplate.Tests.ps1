Describe 'HtmlTemplate' {

    BeforeAll {
        $templatePath = "$PSScriptRoot/../HtmlTemplate.html"
        $template = Get-Content -Path $templatePath -Raw
    }

    It 'file exists' {
        Test-Path $templatePath | Should -Be $true
    }

    It 'is valid HTML with DOCTYPE' {
        $template | Should -BeLike '<!DOCTYPE html>*'
    }

    It 'has UTF-8 charset' {
        $template | Should -BeLike '*charset="UTF-8"*'
    }

    It 'has viewport meta tag' {
        $template | Should -BeLike '*viewport*'
    }

    It 'uses Mermaid v11' {
        $template | Should -BeLike '*mermaid@11*'
    }

    It 'contains VoiceAppNamePlaceHolder for title replacement' {
        $template | Should -BeLike '*VoiceAppNamePlaceHolder*'
    }

    It 'contains VoiceAppNameHtmlIdPlaceHolder for id replacement' {
        $template | Should -BeLike '*VoiceAppNameHtmlIdPlaceHolder*'
    }

    It 'contains ThemePlaceHolder for theme injection' {
        $template | Should -BeLike '*ThemePlaceHolder*'
    }

    It 'contains MermaidPlaceHolder for diagram injection' {
        $template | Should -BeLike '*MermaidPlaceHolder*'
    }

    It 'has mermaid pre tag for diagram rendering' {
        $template | Should -BeLike '*class="mermaid"*'
    }

    It 'has dark mode support via prefers-color-scheme' {
        $template | Should -BeLike '*prefers-color-scheme: dark*'
    }

    It 'links back to the project repository' {
        $template | Should -BeLike '*github.com/mozziemozz/M365CallFlowVisualizer*'
    }

    It 'has no VSCode-specific CSS classes' {
        $template | Should -Not -BeLike '*vscode-dark*'
        $template | Should -Not -BeLike '*vscode-light*'
        $template | Should -Not -BeLike '*vscode-high-contrast*'
    }

    Context 'placeholder replacement works correctly' {

        It 'all placeholders can be replaced' {
            $result = $template -replace 'VoiceAppNamePlaceHolder', 'Test AA' `
                                -replace 'VoiceAppNameHtmlIdPlaceHolder', 'Test-AA' `
                                -replace 'ThemePlaceHolder', '' `
                                -replace 'MermaidPlaceHolder', 'graph TD; A-->B'

            $result | Should -BeLike '*Test AA*'
            $result | Should -BeLike '*Test-AA*'
            $result | Should -BeLike '*graph TD; A-->B*'
            $result | Should -Not -BeLike '*PlaceHolder*'
        }
    }
}
