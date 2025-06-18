# quality_assurance.py - Quality Assurance Agent
import re
import logging
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from fs_processor import FSType

@dataclass
class QualityResult:
    """Container for quality assessment results"""
    score: float
    is_valid: bool
    issues: List[str]
    recommendations: List[str]
    category_scores: Dict[str, float]

class QualityAssuranceAgent:
    """Three-tier quality validation system for financial content"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
        # Quality thresholds
        self.excellent_threshold = 90
        self.good_threshold = 80
        self.acceptable_threshold = 70
        
        # Template artifacts to check
        self.template_artifacts = [
            'Pattern 1:', 'Pattern 2:', 'Pattern 3:',
            '[', ']', '{', '}', 'xxx', 'XXX',
            'template', 'placeholder', 'PLACEHOLDER',
            'TBD', 'TODO', 'FIXME'
        ]
        
        # Professional language indicators
        self.professional_terms = [
            'comprised', 'represented', 'indicated', 'demonstrated',
            'reflected', 'maintained', 'established', 'confirmed',
            'verified', 'assessed', 'evaluated', 'analyzed'
        ]
        
        # Risk language indicators
        self.risk_indicators = [
            'provision', 'impairment', 'restricted', 'covenant',
            'collateral', 'mortgage', 'guarantee', 'contingent'
        ]
    
    def validate_content(self, content: str, fs_type: FSType) -> QualityResult:
        """Comprehensive three-tier content validation"""
        try:
            # Tier 1: Format Compliance
            format_score, format_issues = self._validate_format_compliance(content)
            
            # Tier 2: Data Accuracy
            accuracy_score, accuracy_issues = self._validate_data_accuracy(content, fs_type)
            
            # Tier 3: Content Quality
            quality_score, quality_issues = self._validate_content_quality(content, fs_type)
            
            # Calculate overall score
            overall_score = (format_score * 0.3 + accuracy_score * 0.4 + quality_score * 0.3)
            
            # Compile all issues
            all_issues = format_issues + accuracy_issues + quality_issues
            
            # Generate recommendations
            recommendations = self._generate_recommendations(
                overall_score, all_issues, fs_type
            )
            
            # Category scores
            category_scores = {
                'format_compliance': format_score,
                'data_accuracy': accuracy_score,
                'content_quality': quality_score,
                'overall': overall_score
            }
            
            return QualityResult(
                score=overall_score,
                is_valid=overall_score >= self.acceptable_threshold,
                issues=all_issues,
                recommendations=recommendations,
                category_scores=category_scores
            )
            
        except Exception as e:
            self.logger.error(f"Quality validation failed: {str(e)}")
            return QualityResult(
                score=0.0,
                is_valid=False,
                issues=[f"Validation error: {str(e)}"],
                recommendations=["Manual review required due to validation error"],
                category_scores={}
            )
    
    def _validate_format_compliance(self, content: str) -> Tuple[float, List[str]]:
        """Tier 1: Validate format compliance and template artifact removal"""
        score = 100.0
        issues = []
        
        # Check for template artifacts
        artifact_count = 0
        for artifact in self.template_artifacts:
            if artifact.lower() in content.lower():
                artifact_count += 1
                issues.append(f"Template artifact found: '{artifact}'")
        
        # Penalize for artifacts
        if artifact_count > 0:
            score -= min(50, artifact_count * 10)
        
        # Check for proper markdown structure
        if not re.search(r'^##\s+\w+', content, re.MULTILINE):
            score -= 15
            issues.append("Missing proper markdown headers")
        
        # Check for empty sections
        if re.search(r'###\s+[^\n]+\n\s*\n', content):
            score -= 10
            issues.append("Empty content sections detected")
        
        # Check for proper paragraph structure
        paragraphs = [p.strip() for p in content.split('\n\n') if p.strip()]
        if len(paragraphs) < 3:
            score -= 10
            issues.append("Insufficient content paragraphs")
        
        return max(0, score), issues
    
    def _validate_data_accuracy(self, content: str, fs_type: FSType) -> Tuple[float, List[str]]:
        """Tier 2: Validate data accuracy and completeness"""
        score = 100.0
        issues = []
        
        # Extract currency amounts
        currency_amounts = re.findall(r'(?:CNY|USD|\$)\s*[\d,]+(?:\.\d+)?\s*[KM]?', content)
        
        # Check for missing financial figures
        if len(currency_amounts) < 5:
            score -= 20
            issues.append("Insufficient financial figures in content")
        
        # Check for unrealistic amounts (basic sanity check)
        for amount in currency_amounts:
            if '999999' in amount or '000000' in amount:
                score -= 15
                issues.append(f"Potentially unrealistic amount: {amount}")
        
        # Check for date consistency
        dates = re.findall(r'\d{1,2}/\d{1,2}/\d{4}', content)
        if len(set(dates)) > 2:
            score -= 10
            issues.append("Inconsistent dates found")
        
        # Validate FS-specific requirements
        if fs_type == FSType.BS:
            score, bs_issues = self._validate_bs_accuracy(content, score)
            issues.extend(bs_issues)
        else:
            score, is_issues = self._validate_is_accuracy(content, score)
            issues.extend(is_issues)
        
        return max(0, score), issues
    
    def _validate_bs_accuracy(self, content: str, score: float) -> Tuple[float, List[str]]:
        """Balance Sheet specific accuracy validation"""
        issues = []
        
        # Check for key BS components
        required_sections = ['Assets', 'Liabilities', 'Equity']
        missing_sections = []
        
        for section in required_sections:
            if section.lower() not in content.lower():
                missing_sections.append(section)
        
        if missing_sections:
            score -= len(missing_sections) * 15
            issues.extend([f"Missing {section} section" for section in missing_sections])
        
        # Check for asset-liability balance indicators
        if 'total assets' in content.lower() and 'total liabilities' in content.lower():
            # This is a good sign - shows comprehensive coverage
            pass
        else:
            score -= 10
            issues.append("Missing comprehensive balance sheet coverage")
        
        return score, issues
    
    def _validate_is_accuracy(self, content: str, score: float) -> Tuple[float, List[str]]:
        """Income Statement specific accuracy validation"""
        issues = []
        
        # Check for key IS components
        required_elements = ['revenue', 'cost', 'income']
        missing_elements = []
        
        for element in required_elements:
            if element.lower() not in content.lower():
                missing_elements.append(element)
        
        if missing_elements:
            score -= len(missing_elements) * 15
            issues.extend([f"Missing {element} coverage" for element in missing_elements])
        
        return score, issues
    
    def _validate_content_quality(self, content: str, fs_type: FSType) -> Tuple[float, List[str]]:
        """Tier 3: Validate content quality and professional standards"""
        score = 100.0
        issues = []
        
        # Check professional language usage
        professional_count = sum(1 for term in self.professional_terms 
                                if term.lower() in content.lower())
        
        if professional_count < 3:
            score -= 15
            issues.append("Insufficient professional financial terminology")
        
        # Check for appropriate risk disclosure
        risk_count = sum(1 for term in self.risk_indicators 
                        if term.lower() in content.lower())
        
        # Balance sheet should have some risk considerations
        if fs_type == FSType.BS and risk_count == 0:
            score -= 10
            issues.append("Missing risk factor considerations")
        
        # Check sentence structure and readability
        sentences = re.split(r'[.!?]+', content)
        valid_sentences = [s.strip() for s in sentences if len(s.strip()) > 10]
        
        if len(valid_sentences) < 10:
            score -= 15
            issues.append("Insufficient sentence count for comprehensive analysis")
        
        # Check for overly short sentences (may indicate incomplete content)
        short_sentences = [s for s in valid_sentences if len(s.split()) < 8]
        if len(short_sentences) > len(valid_sentences) * 0.5:
            score -= 10
            issues.append("Too many short sentences - may indicate incomplete analysis")
        
        # Check for repetitive content
        words = content.lower().split()
        unique_words = len(set(words))
        total_words = len(words)
        
        if total_words > 0 and unique_words / total_words < 0.5:
            score -= 15
            issues.append("High repetition detected - content may lack depth")
        
        # Check for balanced coverage across sections
        sections = re.split(r'##\s+', content)[1:]  # Skip first empty split
        if len(sections) < 3:
            score -= 10
            issues.append("Insufficient section coverage")
        
        # Check section length balance
        section_lengths = [len(section.split()) for section in sections]
        if section_lengths:
            avg_length = sum(section_lengths) / len(section_lengths)
            unbalanced_sections = [i for i, length in enumerate(section_lengths) 
                                 if length < avg_length * 0.3]
            
            if len(unbalanced_sections) > 0:
                score -= 5
                issues.append("Unbalanced section lengths detected")
        
        return max(0, score), issues
    
    def _generate_recommendations(
        self, 
        score: float, 
        issues: List[str], 
        fs_type: FSType
    ) -> List[str]:
        """Generate actionable recommendations based on quality assessment"""
        recommendations = []
        
        if score >= self.excellent_threshold:
            recommendations.append("Content meets excellent quality standards")
            
        elif score >= self.good_threshold:
            recommendations.append("Content meets good quality standards with minor improvements needed")
            
        elif score >= self.acceptable_threshold:
            recommendations.append("Content meets acceptable standards but requires attention to identified issues")
            
        else:
            recommendations.append("Content requires significant improvement before publication")
        
        # Specific recommendations based on issues
        issue_categories = {
            'template': [i for i in issues if 'template' in i.lower() or 'artifact' in i.lower()],
            'financial': [i for i in issues if 'amount' in i.lower() or 'figure' in i.lower()],
            'structure': [i for i in issues if 'section' in i.lower() or 'paragraph' in i.lower()],
            'language': [i for i in issues if 'professional' in i.lower() or 'terminology' in i.lower()]
        }
        
        if issue_categories['template']:
            recommendations.append("Remove all template artifacts and placeholders")
        
        if issue_categories['financial']:
            recommendations.append("Verify all financial figures and ensure realistic amounts")
        
        if issue_categories['structure']:
            recommendations.append("Improve content structure and section organization")
        
        if issue_categories['language']:
            recommendations.append("Enhance professional language and financial terminology usage")
        
        # FS-specific recommendations
        if fs_type == FSType.BS:
            recommendations.append("Ensure comprehensive balance sheet coverage including assets, liabilities, and equity")
        else:
            recommendations.append("Ensure comprehensive income statement coverage including revenue, costs, and profitability")
        
        return recommendations
    
    def generate_quality_report(self, result: QualityResult) -> str:
        """Generate a formatted quality assessment report"""
        report_lines = [
            "# Quality Assessment Report",
            "",
            f"**Overall Score:** {result.score:.1f}/100",
            f"**Status:** {'PASS' if result.is_valid else 'FAIL'}",
            "",
            "## Category Scores",
            ""
        ]
        
        for category, score in result.category_scores.items():
            status = "✓" if score >= self.acceptable_threshold else "✗"
            report_lines.append(f"- {category.replace('_', ' ').title()}: {score:.1f} {status}")
        
        if result.issues:
            report_lines.extend([
                "",
                "## Issues Identified",
                ""
            ])
            for i, issue in enumerate(result.issues, 1):
                report_lines.append(f"{i}. {issue}")
        
        if result.recommendations:
            report_lines.extend([
                "",
                "## Recommendations",
                ""
            ])
            for i, rec in enumerate(result.recommendations, 1):
                report_lines.append(f"{i}. {rec}")
        
        return "\n".join(report_lines)