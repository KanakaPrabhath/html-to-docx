#!/usr/bin/env node

/**
 * Main test runner for HTML to DOCX converter
 * Runs all test files in the tests directory
 */

import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath, pathToFileURL } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const testFiles = [
  'test-basic-formatting.js',
  'test-headings.js',
  'test-lists.js',
  'test-tables.js',
  'test-images.js',
  'test-textboxes.js',
  'test-headers-footers.js',
  'test-page-breaks.js',
  'test-page-borders.js',
  'test-page-sizes-fonts.js',
  'test-edge-cases.js'
];

async function runAllTests() {
  console.log('='.repeat(60));
  console.log('HTML to DOCX Converter - Comprehensive Test Suite');
  console.log('='.repeat(60));
  console.log(`Running ${testFiles.length} test files...\n`);

  const results = {
    passed: 0,
    failed: 0,
    total: testFiles.length
  };

  for (const testFile of testFiles) {
    const testPath = path.join(__dirname, testFile);

    try {
      // Check if file exists
      await fs.access(testPath);

      console.log(`Running ${testFile}...`);

      // Import and run the test
      const testModule = await import(pathToFileURL(testPath).href);

      // Assume the test file exports a default function that returns a promise
      if (testModule.default && typeof testModule.default === 'function') {
        const result = await testModule.default();
        if (result !== false) {
          console.log(`✓ ${testFile} passed\n`);
          results.passed++;
        } else {
          console.log(`✗ ${testFile} failed\n`);
          results.failed++;
        }
      } else {
        console.log(`⚠ ${testFile} does not export a default test function\n`);
        results.failed++;
      }

    } catch (error) {
      console.log(`✗ ${testFile} failed to run: ${error.message}\n`);
      results.failed++;
    }
  }

  // Print summary
  console.log('='.repeat(60));
  console.log('Test Results Summary');
  console.log('='.repeat(60));
  console.log(`Total tests: ${results.total}`);
  console.log(`Passed: ${results.passed}`);
  console.log(`Failed: ${results.failed}`);
  console.log(`Success rate: ${((results.passed / results.total) * 100).toFixed(1)}%`);

  if (results.failed === 0) {
    console.log('\n🎉 All tests passed!');
  } else {
    console.log(`\n⚠️  ${results.failed} test(s) failed.`);
  }

  console.log('\nTest output files are saved in the tests/ directory.');
  console.log('You can open the .docx files to verify the results.');

  return results.failed === 0;
}

// Handle command line arguments
const args = process.argv.slice(2);

if (args.includes('--help') || args.includes('-h')) {
  console.log('HTML to DOCX Test Runner');
  console.log('');
  console.log('Usage: node run-tests.js [options]');
  console.log('');
  console.log('Options:');
  console.log('  --help, -h    Show this help message');
  console.log('  --list        List all available test files');
  console.log('  --run <file>  Run a specific test file');
  console.log('');
  console.log('Examples:');
  console.log('  node run-tests.js              # Run all tests');
  console.log('  node run-tests.js --list       # List test files');
  console.log('  node run-tests.js --run test-basic-formatting.js');
  process.exit(0);
}

if (args.includes('--list')) {
  console.log('Available test files:');
  testFiles.forEach(file => console.log(`  - ${file}`));
  process.exit(0);
}

if (args.includes('--run')) {
  const fileIndex = args.indexOf('--run');
  if (fileIndex + 1 < args.length) {
    const specificFile = args[fileIndex + 1];
    if (testFiles.includes(specificFile)) {
      console.log(`Running specific test: ${specificFile}`);
      // Run only the specified test
      const testPath = path.join(__dirname, specificFile);
      import(pathToFileURL(testPath).href).then(module => {
        if (module.default) {
          module.default().catch(console.error);
        } else {
          console.error(`Test file ${specificFile} does not export a default function`);
        }
      }).catch(console.error);
    } else {
      console.error(`Test file not found: ${specificFile}`);
      console.log('Available test files:');
      testFiles.forEach(file => console.log(`  - ${file}`));
      process.exit(1);
    }
  } else {
    console.error('Please specify a test file to run');
    process.exit(1);
  }
} else {
  // Run all tests
  runAllTests().then(success => {
    process.exit(success ? 0 : 1);
  }).catch(error => {
    console.error('Test runner failed:', error);
    process.exit(1);
  });
}