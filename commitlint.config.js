module.exports = {
    extends: ['@commitlint/config-conventional'],
    rules: {
        'type-enum': [
            2,
            'always',
            ['feat', 'fix', 'docs', 'style', 'refactor', 'test', 'revert', 'perf', "ci", "build"],
        ],
        // Disable body/footer line length limits so trailer lines such as
        // `Agent-Logs-Url: <long-url>` (auto-appended by automation) do not
        // block otherwise-valid conventional commits.
        'body-max-line-length': [0, 'always'],
        'footer-max-line-length': [0, 'always'],
    },
};