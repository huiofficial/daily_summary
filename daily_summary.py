#!/usr/bin/env python3
"""
Daily Git Commit Summary Tool
Summarizes git commits from multiple repositories into an Excel report.
"""

import os
import sys
import subprocess
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Tuple
import pandas as pd

try:
    import ollama
    OLLAMA_AVAILABLE = True
except ImportError:
    OLLAMA_AVAILABLE = False
    print("Warning: ollama library not found. Install with: pip install ollama")


def find_git_repos(base_dir: str) -> List[str]:
    """Find all git repositories in the given directory."""
    repos = []
    base_path = Path(base_dir)
    
    if not base_path.exists():
        print(f"Error: Directory '{base_dir}' does not exist.")
        return repos
    
    # Walk through directory to find .git folders
    for root, dirs, files in os.walk(base_path):
        # Skip hidden directories and common non-repo directories
        dirs[:] = [d for d in dirs if not d.startswith('.') and d not in ['node_modules', 'venv', '__pycache__']]
        
        git_dir = Path(root) / '.git'
        if git_dir.exists() and git_dir.is_dir():
            repos.append(str(root))
    
    return repos


def get_commits(repo_path: str, start_date: datetime, end_date: datetime, author: str = None) -> List[Dict]:
    """Get commits from a repository within the specified date range."""
    commits = []
    
    try:
        # Format dates for git log
        date_format = "%Y-%m-%d"
        start_str = start_date.strftime(date_format)
        end_str = (end_date + timedelta(days=1)).strftime(date_format)  # Include end date
        
        # Build git log command
        cmd = [
            'git', 'log',
            '--since', start_str,
            '--until', end_str,
            '--pretty=format:%H|%an|%ae|%ad|%s',
            '--date=format:%Y-%m-%d %H:%M:%S'
        ]
        
        if author:
            cmd.extend(['--author', author])
        
        # Change to repo directory and execute
        result = subprocess.run(
            cmd,
            cwd=repo_path,
            capture_output=True,
            text=True,
            timeout=30
        )
        
        if result.returncode != 0:
            if result.stderr:
                print(f"Warning: Error getting commits from {repo_path}: {result.stderr.strip()}")
            return commits
        
        # Parse commit output
        for line in result.stdout.strip().split('\n'):
            if not line:
                continue
            
            parts = line.split('|', 4)
            if len(parts) == 5:
                commit_hash, author_name, author_email, commit_date, message = parts
                
                # Get file changes count
                stats_cmd = ['git', 'show', '--stat', '--format=', commit_hash]
                stats_result = subprocess.run(
                    stats_cmd,
                    cwd=repo_path,
                    capture_output=True,
                    text=True,
                    timeout=10
                )
                
                files_changed = 0
                insertions = 0
                deletions = 0
                
                if stats_result.returncode == 0:
                    for stat_line in stats_result.stdout.strip().split('\n'):
                        if 'files changed' in stat_line:
                            # Parse stats: " 3 files changed, 45 insertions(+), 12 deletions(-)"
                            parts = stat_line.split(',')
                            if len(parts) > 0:
                                try:
                                    files_changed = int(parts[0].strip().split()[0])
                                except:
                                    pass
                            if len(parts) > 1:
                                if 'insertion' in parts[1]:
                                    try:
                                        insertions = int(parts[1].strip().split()[0])
                                    except:
                                        pass
                            if len(parts) > 2:
                                if 'deletion' in parts[2]:
                                    try:
                                        deletions = int(parts[2].strip().split()[0])
                                    except:
                                        pass
                
                commits.append({
                    'repository': os.path.basename(repo_path),
                    'commit_hash': commit_hash[:8],  # Short hash
                    'author': author_name,
                    'email': author_email,
                    'date': commit_date,
                    'message': message,
                    'files_changed': files_changed,
                    'insertions': insertions,
                    'deletions': deletions,
                })
    
    except subprocess.TimeoutExpired:
        print(f"Warning: Timeout getting commits from {repo_path}")
    except Exception as e:
        print(f"Warning: Error processing {repo_path}: {str(e)}")
    
    return commits


def get_repo_info(repo_path: str) -> Dict:
    """Get repository information like remote URL."""
    info = {
        'remote_url': 'N/A',
        'branch': 'N/A'
    }
    
    try:
        # Get remote URL
        result = subprocess.run(
            ['git', 'remote', 'get-url', 'origin'],
            cwd=repo_path,
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            info['remote_url'] = result.stdout.strip()
        
        # Get current branch
        result = subprocess.run(
            ['git', 'branch', '--show-current'],
            cwd=repo_path,
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            info['branch'] = result.stdout.strip()
    except:
        pass
    
    return info


def summarize_repo_commits(repo_name: str, commit_messages: List[str], model: str = 'qwen3:0.6b') -> str:
    """Use ollama to summarize commit messages for a repository."""
    if not OLLAMA_AVAILABLE:
        return "Ollama not available. Install with: pip install ollama"
    
    if not commit_messages:
        return "No commits to summarize."
    
    try:
        # Combine commit messages (limit to avoid token limit)
        # Take at most the first 50 commit messages to avoid overwhelming the model
        messages_to_summarize = commit_messages[:50]
        messages_text = "\n".join([f"{i+1}. {msg}" for i, msg in enumerate(messages_to_summarize)])
        
        if len(commit_messages) > 50:
            messages_text += f"\n... (还有 {len(commit_messages) - 50} 条提交信息未显示)"
        
        # Create prompt for summarization
        prompt = f"""请总结以下项目 '{repo_name}' 的提交信息，生成一个简洁的工作总结（用中文）：

提交信息列表：
{messages_text}

请用2-5句话总结这个项目在这个时间段的主要工作内容，要求简洁明了，突出主要功能和改进点。"""

        # Call ollama API
        response = ollama.chat(
            model=model,
            messages=[
                {
                    'role': 'user',
                    'content': prompt
                }
            ]
        )
        
        summary = response['message']['content'].strip()
        return summary
    
    except Exception as e:
        error_msg = f"Error generating summary: {str(e)}"
        print(f"    ⚠️  {error_msg}")
        return error_msg


def generate_excel_report(commits: List[Dict], output_file: str, period: str, repo_summaries: Dict[str, str] = None):
    """Generate Excel report from commits."""
    if not commits:
        print("No commits found for the specified period.")
        return
    
    # Create DataFrame
    df = pd.DataFrame(commits)
    
    # Reorder columns
    columns_order = ['repository', 'date', 'author', 'email', 'commit_hash', 'message', 
                     'files_changed', 'insertions', 'deletions']
    df = df[columns_order]
    
    # Sort by date (newest first)
    df = df.sort_values('date', ascending=False)
    
    # Create Excel writer with formatting
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Commits', index=False)
        
        # Get workbook and worksheet for formatting
        workbook = writer.book
        worksheet = writer.sheets['Commits']
        
        # Auto-adjust column widths
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Alignment, Font
        
        for idx, col in enumerate(df.columns, 1):
            max_length = max(
                df[col].astype(str).map(len).max(),
                len(col)
            ) + 2
            worksheet.column_dimensions[get_column_letter(idx)].width = min(max_length, 50)
        
        # Freeze header row
        worksheet.freeze_panes = 'A2'
        
        # Add summary sheet
        summary_data = {
            'Metric': [
                'Period',
                'Total Commits',
                'Total Repositories',
                'Total Files Changed',
                'Total Insertions',
                'Total Deletions',
                'Unique Authors'
            ],
            'Value': [
                period,
                len(commits),
                df['repository'].nunique(),
                df['files_changed'].sum(),
                df['insertions'].sum(),
                df['deletions'].sum(),
                df['author'].nunique()
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Format summary sheet
        summary_ws = writer.sheets['Summary']
        for idx in range(1, len(summary_data['Metric']) + 1):
            summary_ws.column_dimensions[get_column_letter(idx)].width = 25
        
        # Add repository summaries sheet if available
        if repo_summaries:
            summaries_data = []
            for repo_name, summary in repo_summaries.items():
                summaries_data.append({
                    'Repository': repo_name,
                    'Summary': summary
                })
            
            summaries_df = pd.DataFrame(summaries_data)
            summaries_df.to_excel(writer, sheet_name='Repository Summaries', index=False)
            
            # Format repository summaries sheet
            summaries_ws = writer.sheets['Repository Summaries']
            summaries_ws.column_dimensions['A'].width = 30
            summaries_ws.column_dimensions['B'].width = 80
            
            # Enable text wrapping for summary column
            for row in range(2, len(summaries_df) + 2):
                cell = summaries_ws[f'B{row}']
                cell.alignment = Alignment(wrap_text=True, vertical='top')
    
    print(f"\n{'='*70}")
    print(f"Excel report generated: {output_file}")
    print(f"{'='*70}")
    print(f"Total commits: {len(commits)}")
    print(f"Total repositories: {df['repository'].nunique()}")
    if repo_summaries:
        print(f"Repository summaries generated: {len(repo_summaries)}")
        print(f"\nThe Excel file contains the following sheets:")
        print(f"  1. Commits - Detailed commit information")
        print(f"  2. Summary - Aggregated statistics")
        print(f"  3. Repository Summaries - AI-generated summaries for each repository")
    print(f"{'='*70}\n")


def main():
    parser = argparse.ArgumentParser(
        description='Summarize daily git commits from multiple repositories into Excel report',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Summarize today's commits
  python daily_summary.py /path/to/work/dir --today

  # Summarize yesterday's commits
  python daily_summary.py /path/to/work/dir --yesterday

  # Summarize last week's commits
  python daily_summary.py /path/to/work/dir --lastweek

  # Custom date range
  python daily_summary.py /path/to/work/dir --start 2024-01-01 --end 2024-01-07
        """
    )
    
    parser.add_argument('work_dir', help='Directory containing git repositories')
    parser.add_argument('--today', action='store_true', help='Summarize today\'s commits')
    parser.add_argument('--yesterday', action='store_true', help='Summarize yesterday\'s commits')
    parser.add_argument('--lastweek', action='store_true', help='Summarize last week\'s commits (7 days)')
    parser.add_argument('--start', type=str, help='Start date (YYYY-MM-DD)')
    parser.add_argument('--end', type=str, help='End date (YYYY-MM-DD)')
    parser.add_argument('--output', '-o', type=str, help='Output Excel file path (default: auto-generated)')
    parser.add_argument('--author', type=str, help='Filter commits by author name/email')
    
    args = parser.parse_args()
    
    # Determine date range
    now = datetime.now()
    start_date = None
    end_date = None
    period_name = ""
    
    if args.today:
        start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = now
        period_name = "Today"
    elif args.yesterday:
        yesterday = now - timedelta(days=1)
        start_date = yesterday.replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = yesterday.replace(hour=23, minute=59, second=59, microsecond=999999)
        period_name = "Yesterday"
    elif args.lastweek:
        end_date = now
        start_date = (now - timedelta(days=7)).replace(hour=0, minute=0, second=0, microsecond=0)
        period_name = "Last Week (7 days)"
    elif args.start and args.end:
        try:
            start_date = datetime.strptime(args.start, '%Y-%m-%d')
            end_date = datetime.strptime(args.end, '%Y-%m-%d')
            period_name = f"{args.start} to {args.end}"
        except ValueError:
            print("Error: Invalid date format. Use YYYY-MM-DD")
            sys.exit(1)
    else:
        # Default to today if no option specified
        start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = now
        period_name = "Today"
    
    # Generate output filename if not provided
    if not args.output:
        date_str = start_date.strftime('%Y%m%d')
        if args.yesterday or args.lastweek:
            date_str = end_date.strftime('%Y%m%d')
        args.output = f"commit_summary_{date_str}.xlsx"
    
    # Find all git repositories
    print(f"Scanning for git repositories in: {args.work_dir}")
    repos = find_git_repos(args.work_dir)
    
    if not repos:
        print("No git repositories found in the specified directory.")
        sys.exit(1)
    
    print(f"Found {len(repos)} git repositories")
    
    # Collect commits from all repositories
    all_commits = []
    repo_commits_map = {}  # Map repo name to list of commits
    
    for repo in repos:
        repo_name = os.path.basename(repo)
        print(f"Processing: {repo_name}")
        commits = get_commits(repo, start_date, end_date, args.author)
        all_commits.extend(commits)
        repo_commits_map[repo_name] = commits
    
    # Generate summaries for each repository using ollama
    repo_summaries = {}
    if OLLAMA_AVAILABLE:
        print("\n" + "="*70)
        print("Generating repository summaries using ollama (qwen3:0.6b)...")
        print("="*70)
        for repo_name, commits in repo_commits_map.items():
            if commits:
                print(f"\n[{repo_name}]")
                print(f"  Commits: {len(commits)}")
                commit_messages = [commit['message'] for commit in commits]
                print(f"  Generating summary...")
                summary = summarize_repo_commits(repo_name, commit_messages)
                repo_summaries[repo_name] = summary
                print(f"  Summary: {summary}")
        print("\n" + "="*70)
        print("Summary generation completed.")
        print("="*70 + "\n")
    else:
        print("\nWarning: Ollama not available. Skipping repository summaries.")
        print("Install ollama: pip install ollama")
    
    # Generate Excel report
    generate_excel_report(all_commits, args.output, period_name, repo_summaries)


if __name__ == '__main__':
    main()

