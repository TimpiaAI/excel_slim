use std::path::PathBuf;

use clap::{Parser, ValueEnum};
use excel_slim_core::{analyze_path, optimize_path, MediaMode, Options, Profile, VbaMode};

#[derive(Debug, Clone, ValueEnum)]
enum ProfileCli {
    Safe,
    Balanced,
    Aggressive,
}

impl From<ProfileCli> for Profile {
    fn from(value: ProfileCli) -> Self {
        match value {
            ProfileCli::Safe => Profile::Safe,
            ProfileCli::Balanced => Profile::Balanced,
            ProfileCli::Aggressive => Profile::Aggressive,
        }
    }
}

#[derive(Debug, Clone, ValueEnum)]
enum VbaCli {
    Auto,
    Off,
    On,
}

impl From<VbaCli> for VbaMode {
    fn from(value: VbaCli) -> Self {
        match value {
            VbaCli::Auto => VbaMode::Auto,
            VbaCli::Off => VbaMode::Off,
            VbaCli::On => VbaMode::On,
        }
    }
}

#[derive(Debug, Clone, ValueEnum)]
enum MediaCli {
    Off,
    Lossless,
    Lossy,
}

impl From<MediaCli> for MediaMode {
    fn from(value: MediaCli) -> Self {
        match value {
            MediaCli::Off => MediaMode::Off,
            MediaCli::Lossless => MediaMode::Lossless,
            MediaCli::Lossy => MediaMode::Lossy,
        }
    }
}

#[derive(Debug, Clone, ValueEnum)]
enum ReportFormat {
    Human,
    Json,
}

#[derive(Parser, Debug)]
#[command(name = "excel-slim", version, about = "Slim Excel workbooks safely")]
struct Cli {
    input: PathBuf,

    #[arg(long)]
    output: Option<PathBuf>,

    #[arg(long, value_enum, default_value_t = ProfileCli::Safe)]
    profile: ProfileCli,

    #[arg(long, default_value_t = true)]
    xml: bool,

    #[arg(long, default_value_t = true)]
    zip: bool,

    #[arg(long, value_enum, default_value_t = VbaCli::Auto)]
    vba: VbaCli,

    #[arg(long, value_enum, default_value_t = MediaCli::Off)]
    media: MediaCli,

    #[arg(
        long,
        value_enum,
        default_value = "human",
        default_missing_value = "human",
        num_args = 0..=1
    )]
    report: ReportFormat,

    #[arg(long)]
    analyze: bool,

    #[arg(long)]
    auto: bool,
}

fn main() {
    let cli = Cli::parse();

    if cli.analyze {
        match analyze_path(&cli.input) {
            Ok(report) => match cli.report {
                ReportFormat::Json => {
                    let json = serde_json::to_string_pretty(&report).unwrap_or_default();
                    println!("{json}");
                }
                ReportFormat::Human => print_analysis(&report),
            },
            Err(err) => {
                eprintln!("{err}");
                std::process::exit(1);
            }
        }
        return;
    }

    let mut options = Options {
        profile: cli.profile.into(),
        xml: cli.xml,
        zip: cli.zip,
        vba: cli.vba.into(),
        media: cli.media.into(),
    };

    if cli.auto {
        options = Options::default();
    }

    match optimize_path(&cli.input, cli.output.as_deref(), options) {
        Ok(report) => match cli.report {
            ReportFormat::Json => {
                let json = serde_json::to_string_pretty(&report).unwrap_or_default();
                println!("{json}");
            }
            ReportFormat::Human => print_report(&report),
        },
        Err(err) => {
            eprintln!("{err}");
            std::process::exit(1);
        }
    }
}

fn print_report(report: &excel_slim_core::OptimizationReport) {
    println!("Input: {}", report.input_path);
    println!("Output: {}", report.output_path);
    println!("Format: {}", report.format);
    println!("Original size: {} bytes", report.original_size_bytes);
    println!("Final size: {} bytes", report.final_size_bytes);
    println!(
        "Delta: {} bytes ({:.2}%)",
        report.delta_bytes, report.delta_percent
    );

    if !report.modules.is_empty() {
        println!("Modules:");
        for module in &report.modules {
            println!(
                "- {}: {} -> {} bytes ({:.2}%)",
                module.name, module.bytes_before, module.bytes_after, module.delta_percent
            );
        }
    }

    if !report.notes.is_empty() {
        println!("Notes:");
        for note in &report.notes {
            println!("- {note}");
        }
    }

    if !report.warnings.is_empty() {
        println!("Warnings:");
        for warning in &report.warnings {
            println!("- {warning}");
        }
    }
}

fn print_analysis(report: &excel_slim_core::AnalysisReport) {
    println!("Format: {}", report.format);
    println!("Size: {} bytes", report.size_bytes);
    println!("Has VBA: {}", report.has_vba);
    println!("Has media: {}", report.has_media);
    println!("Worksheets: {}", report.xml_stats.worksheets);
    println!("Shared strings size: {} bytes", report.xml_stats.shared_strings_bytes);
    println!("Styles size: {} bytes", report.xml_stats.styles_bytes);

    if !report.recommendations.is_empty() {
        println!("Recommendations:");
        for rec in &report.recommendations {
            println!("- {rec}");
        }
    }

    if !report.risks.is_empty() {
        println!("Risks:");
        for risk in &report.risks {
            println!("- {risk}");
        }
    }
}
