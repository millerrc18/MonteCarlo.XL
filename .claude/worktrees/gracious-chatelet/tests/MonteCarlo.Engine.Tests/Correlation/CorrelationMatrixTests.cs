using FluentAssertions;
using MonteCarlo.Engine.Correlation;
using Xunit;

namespace MonteCarlo.Engine.Tests.Correlation;

public class CorrelationMatrixTests
{
    [Fact]
    public void Identity_IsValid()
    {
        var matrix = CorrelationMatrix.Identity(3);
        matrix.Size.Should().Be(3);
        matrix.Validate(); // Should not throw
    }

    [Fact]
    public void ValidSymmetricMatrix_PassesValidation()
    {
        var data = new double[,]
        {
            { 1.0, 0.5, 0.3 },
            { 0.5, 1.0, -0.2 },
            { 0.3, -0.2, 1.0 }
        };
        var matrix = new CorrelationMatrix(data);
        matrix.Validate(); // Should not throw
    }

    [Fact]
    public void NonSymmetric_ThrowsOnValidate()
    {
        var data = new double[,]
        {
            { 1.0, 0.5 },
            { 0.3, 1.0 }
        };
        var matrix = new CorrelationMatrix(data);
        var act = () => matrix.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*symmetric*");
    }

    [Fact]
    public void DiagonalNotOne_ThrowsOnValidate()
    {
        var data = new double[,]
        {
            { 0.9, 0.5 },
            { 0.5, 1.0 }
        };
        var matrix = new CorrelationMatrix(data);
        var act = () => matrix.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*Diagonal*");
    }

    [Fact]
    public void ValueOutOfRange_ThrowsOnValidate()
    {
        var data = new double[,]
        {
            { 1.0, 1.5 },
            { 1.5, 1.0 }
        };
        var matrix = new CorrelationMatrix(data);
        var act = () => matrix.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*outside*");
    }

    [Fact]
    public void NonPSD_ThrowsOnValidate()
    {
        // This matrix is symmetric but not PSD
        var data = new double[,]
        {
            { 1.0, 0.9, 0.9 },
            { 0.9, 1.0, -0.9 },
            { 0.9, -0.9, 1.0 }
        };
        var matrix = new CorrelationMatrix(data);
        var act = () => matrix.Validate();
        act.Should().Throw<ArgumentException>().WithMessage("*positive semi-definite*");
    }

    [Fact]
    public void NonPSD_EnsurePSD_ProducesValidMatrix()
    {
        var data = new double[,]
        {
            { 1.0, 0.9, 0.9 },
            { 0.9, 1.0, -0.9 },
            { 0.9, -0.9, 1.0 }
        };
        var matrix = new CorrelationMatrix(data);

        var corrected = matrix.EnsurePositiveSemiDefinite();

        corrected.IsPositiveSemiDefinite().Should().BeTrue();
        // Diagonal should be 1.0
        for (int i = 0; i < corrected.Size; i++)
            corrected[i, i].Should().BeApproximately(1.0, 1e-10);
        // Should be symmetric
        for (int i = 0; i < corrected.Size; i++)
            for (int j = i + 1; j < corrected.Size; j++)
                corrected[i, j].Should().BeApproximately(corrected[j, i], 1e-10);
    }

    [Fact]
    public void AlreadyPSD_EnsurePSD_ReturnsSame()
    {
        var data = new double[,]
        {
            { 1.0, 0.5 },
            { 0.5, 1.0 }
        };
        var matrix = new CorrelationMatrix(data);

        var result = matrix.EnsurePositiveSemiDefinite();

        // Should return the same object since it's already PSD
        result.Should().BeSameAs(matrix);
    }

    [Fact]
    public void ToArray_ReturnsCopy()
    {
        var matrix = CorrelationMatrix.Identity(2);
        var arr = matrix.ToArray();
        arr[0, 1] = 0.99;
        // Original should be unchanged
        matrix[0, 1].Should().Be(0.0);
    }

    [Fact]
    public void Constructor_NonSquare_Throws()
    {
        var act = () => new CorrelationMatrix(new double[2, 3]);
        act.Should().Throw<ArgumentException>().WithMessage("*square*");
    }

    [Fact]
    public void Constructor_TooSmall_Throws()
    {
        var act = () => new CorrelationMatrix(new double[1, 1]);
        act.Should().Throw<ArgumentException>().WithMessage("*2×2*");
    }

    [Fact]
    public void Identity_TooSmall_Throws()
    {
        var act = () => CorrelationMatrix.Identity(1);
        act.Should().Throw<ArgumentException>();
    }
}
