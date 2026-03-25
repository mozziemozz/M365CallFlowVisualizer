BeforeAll {
    . "$PSScriptRoot/../Functions/Optimize-DisplayName.ps1"
}

Describe 'Optimize-DisplayName' {

    It 'removes parentheses' {
        Optimize-DisplayName -String 'Hello (World)' | Should -Be 'Hello World'
    }

    It 'removes square brackets' {
        Optimize-DisplayName -String 'Test [Group]' | Should -Be 'Test Group'
    }

    It 'removes pipe characters' {
        Optimize-DisplayName -String 'Sales|Support' | Should -Be 'SalesSupport'
    }

    It 'removes tildes' {
        Optimize-DisplayName -String 'Test~Name' | Should -Be 'TestName'
    }

    It 'collapses double spaces' {
        Optimize-DisplayName -String 'Hello  World' | Should -Be 'Hello World'
    }

    It 'replaces @ with at' {
        Optimize-DisplayName -String 'user@domain.com' | Should -Be 'user at domain.com'
    }

    It 'capitalizes Call in call' {
        Optimize-DisplayName -String 'main call queue' | Should -Be 'main Call queue'
    }

    It 'replaces smart quotes with straight quotes' {
        $smartQuote = [char]0x2019
        Optimize-DisplayName -String "it${smartQuote}s working" | Should -Be "it's working"
    }

    It 'trims whitespace' {
        Optimize-DisplayName -String '  Hello World  ' | Should -Be 'Hello World'
    }

    It 'handles empty-ish input after stripping' {
        Optimize-DisplayName -String '()[]|~' | Should -Be ''
    }

    It 'handles combined special characters' {
        Optimize-DisplayName -String 'Sales (Main) [US] | Support~Team' | Should -Be 'Sales Main US SupportTeam'
    }

    It 'handles normal text without changes' {
        Optimize-DisplayName -String 'Regular Name' | Should -Be 'Regular Name'
    }
}
