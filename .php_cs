<?php

$finder = Symfony\CS\Finder\DefaultFinder::create()
    ->in(__DIR__ . '/lib')
    ->in(__DIR__ . '/tests')
;

$fixers = array(
    'concat_with_spaces',
    'elseif',
    'empty_return',
    'encoding',
    'eof_ending',
    'ereg_to_preg',
    'extra_empty_lines',
    'function_declaration',
    'include',
    'indentation',
    'internal-tab',
    'join_function',
    'line_after_namespace',
    'linefeed',
    'lowercase_constants',
    'lowercase_keywords',
    'multiple_use',
    'namespace_no_leading_whitespace',
    'negation-spacer',
    'new_with_braces',
    'object_operator',
    'operators_spaces',
    'use-order',
    'parenthesis',
    'php_closing_tag',
    'phpdoc_indent',
    'phpdoc_no_empty_return',
    'phpdoc_no_package',
    'phpdoc_order',
    'phpdoc_params',
    'phpdoc_separation',
    'phpdoc_trim',
    // 'phpdoc_type_to_var',
    // 'phpdoc_var_to_type',
    'phpdoc_var_without_name',
    'psr0',
    'remove_leading_slash_use',
    'remove_lines_between_uses',
    'return',
    'short_tag',
    'single_blank_line_before_namespace',
    'single_line_after_imports',
    'spaces_before_semicolon',
    'spaces_cast',
    'standardize_not_equal',
    // 'strict',
    'ternary_spaces',
    'trailing_spaces',
    'unused_use',
    'utf8',
    'visibility',
    'whitespacy_lines',
);

return Symfony\CS\Config\Config::create()
    ->setUsingCache(true)
    ->level(null)
    ->finder($finder)
    ->fixers($fixers)
;
