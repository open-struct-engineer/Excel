Attribute VB_Name = "CantileverLeftFormulas"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' BSD 3-Clause License                                                            '''
'''                                                                                 '''
''' Copyright (c) 2020, open-struct-engineer                                        '''
''' All rights reserved.                                                            '''
'''                                                                                 '''
''' Redistribution and use in source and binary forms, with or without              '''
''' modification, are permitted provided that the following conditions are met:     '''
'''                                                                                 '''
''' 1. Redistributions of source code must retain the above copyright notice, this  '''
'''   list of conditions and the following disclaimer.                              '''
'''                                                                                 '''
''' 2. Redistributions in binary form must reproduce the above copyright notice,    '''
'''   this list of conditions and the following disclaimer in the documentation     '''
'''   and/or other materials provided with the distribution.                        '''
'''                                                                                 '''
''' 3. Neither the name of the copyright holder nor the names of its                '''
'''   contributors may be used to endorse or promote products derived from          '''
'''   this software without specific prior written permission.                      '''
'''                                                                                 '''
''' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"     '''
''' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE       '''
''' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE  '''
''' DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE    '''
''' FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL      '''
''' DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR      '''
''' SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER      '''
''' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,   '''
''' OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE   '''
''' OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.            '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function cant_left_initialSlope(slope As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Starting slope on a left side cantilever                 '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at left Support (clockwise positive)               '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'left End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Common Calculations   ''
    ''''''''''''''''''''''''''''
    
    'Support Reactions
    rl = 0
    rr = 0
    
    'Integration Constants
    c1 = slope
    c2 = -1 * c1 * L
    
    '''''''''''''''''''''''''''
    ''  Result Selection     ''
    '''''''''''''''''''''''''''
    
    If result = 0 Then
        
        'Left Support Reaction
        cant_left_initialSlope = rl
        
    ElseIf result = 1 Then
        
        'left Support Reaction
        cant_left_initialSlope = rr
    
    ElseIf result = 2 Then
        
        'Shear at x
        vx = 0
  
        cant_left_initialSlope = vx
    
    ElseIf result = 3 Then
    
        'Moment at x
        mx = 0
        
        cant_left_initialSlope = mx
            
    ElseIf result = 4 Then
        
        'Cross Section Rotation/Slope at x
        If 0 <= x And x <= L Then
            sx = c1
        Else
            sx = 0
        End If
        
        cant_left_initialSlope = sx
    
    ElseIf result = 5 Then
        
        'Deflection at x
        If 0 <= x And x <= L Then
            dx = (c1 * x) + c2
        Else
            dx = 0
        End If
        
        cant_left_initialSlope = dx
    
    ElseIf result = 6 Then
        'Fixed End Moment Left
        femL = 0
        
        cant_left_initialSlope = femL
    
    ElseIf result = 7 Then
    
        'Fixed End Moment left
        femR = 0
        
        cant_left_initialSlope = femR
    
    Else
    
        cant_left_initialSlope = CVErr(xlErrNA)
            
    End If
    
End Function

Function cant_left_point_load(p As Double, a As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Point Load anywhere on a left side cantilever           '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at left Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim b As Double     'Distance from load to free end
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'left End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If p = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        cant_left_point_load = 0
    Else
    
        If a > L Or a < 0 Then
            cant_left_point_load = " A , must be between 0 and L"
        
        Else
          ''''''''''''''''''''''''''''
          ''  Common Calculations   ''
          ''''''''''''''''''''''''''''
          
          b = L - a
          
          'Support Reactions
          rl = 0
          rr = p
          ml = 0
          mr = -1 * p * (L - a)
          
          'Integration Constants
          c3 = 0 + (0.5 * p * (L - a) * (L - a))
          c4 = ((1 / 6) * p * (L - a) * (L - a) * (L - a)) - (c3 * L)
          c1 = c3
          c2 = (c3 * a) + c4 - (c1 * a)
          
          '''''''''''''''''''''''''''
          ''  Result Selection     ''
          '''''''''''''''''''''''''''
          
          If result = 0 Then
              
              'Left Support Reaction
              cant_left_point_load = rl
              
          ElseIf result = 1 Then
              
              'left Support Reaction
              cant_left_point_load = rr
          
          ElseIf result = 2 Then
              
              'Shear at x
              If 0 <= x And x <= a Then
                  vx = 0
              ElseIf a < x And x <= L Then
                  vx = -1 * p
              Else
                  vx = 0
              End If
        
              cant_left_point_load = vx
          
          ElseIf result = 3 Then
          
              'Moment at x
              If 0 <= x And x <= a Then
                  mx = 0
              ElseIf a < x And x <= L Then
                  mx = -1 * p * (x - a)
              Else
                  mx = 0
              End If
              
              cant_left_point_load = mx
                  
          ElseIf result = 4 Then
              
              'Cross Section Rotation/Slope at x
              If 0 <= x And x <= a Then
                  sx = c1 / (E * I)
              ElseIf a < x And x <= L Then
                  sx = ((-0.5 * p * (x - a) * (x - a)) + c3) / (E * I)
              Else
                  sx = 0
              End If
              
              cant_left_point_load = sx
          
          ElseIf result = 5 Then
              
              'Deflection at x
              If 0 <= x And x <= a Then
                  dx = ((c1 * x) + c2) / (E * I)
              ElseIf a < x And x <= L Then
                  dx = (((-1 / 6) * p * (x - a) * (x - a) * (x - a)) + (c3 * x) + c4) / (E * I)
              Else
                  dx = 0
              End If
              
              cant_left_point_load = dx
          
          ElseIf result = 6 Then
              'Fixed End Moment Left
              femL = 0
              
              cant_left_point_load = femL
          
          ElseIf result = 7 Then
          
              'Fixed End Moment left
              femR = mr
              
              cant_left_point_load = femR
          
          Else
          
              cant_left_point_load = CVErr(xlErrNA)
                  
          End If
        End If
    End If
    
End Function

Function cant_left_point_moment(m As Double, a As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Point Moment anywhere on a left side cantilever         '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at left Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim b As Double     'Distance from load to free end
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'left End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If m = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        cant_left_point_moment = 0
    Else
    
        If a > L Or a < 0 Then
            cant_left_point_moment = " A , must be between 0 and L"
        
        Else
            ''''''''''''''''''''''''''''
            ''  Common Calculations   ''
            ''''''''''''''''''''''''''''
            
            b = L - a
            
            'Support Reactions
            rl = 0
            rr = 0
            mr = m
            
            'Integration Constants
            c3 = 0 - (m * L)
            c4 = (-0.5 * m * L * L) - (c3 * L)
            c1 = (1 * m * a) + c3
            c2 = (0.5 * m * a * a) + (c3 * a) + c4 - (c1 * a)
            
            '''''''''''''''''''''''''''
            ''  Result Selection     ''
            '''''''''''''''''''''''''''
            
            If result = 0 Then
                
                'Left Support Reaction
                cant_left_point_moment = rl
                
            ElseIf result = 1 Then
                
                'left Support Reaction
                cant_left_point_moment = rr
            
            ElseIf result = 2 Then
                
                'Shear at x
                vx = 0
            
                cant_left_point_moment = vx
            
            ElseIf result = 3 Then
            
                'Moment at x
                If 0 <= x And x <= a Then
                    mx = 0
                
                ElseIf a < x And x <= L Then
                    mx = m
                Else
                    mx = 0
                End If
                
                cant_left_point_moment = mx
                    
            ElseIf result = 4 Then
                
                'Cross Section Rotation/Slope at x
                If 0 <= x And x <= a Then
                    sx = c1 / (E * I)
                ElseIf a < x And x <= L Then
                    sx = ((m * x) + c3) / (E * I)
                Else
                    sx = 0
                End If
                
                cant_left_point_moment = sx
            
            ElseIf result = 5 Then
                
                'Deflection at x
                If 0 <= x And x <= a Then
                    dx = ((c1 * x) + c2) / (E * I)
                ElseIf a < x And x <= L Then
                    dx = ((0.5 * m * x * x) + (c3 * x) + c4) / (E * I)
                Else
                    dx = 0
                End If
                
                cant_left_point_moment = dx
            
            ElseIf result = 6 Then
                'Fixed End Moment Left
                femL = 0
                
                cant_left_point_moment = femL
            
            ElseIf result = 7 Then
            
                'Fixed End Moment left
                femR = mr
                
                cant_left_point_moment = femR
            
            Else
            
                cant_left_point_moment = CVErr(xlErrNA)
                    
            End If
        End If
    End If
    
End Function

Function cant_left_uniform_load(w As Double, a As Double, b As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Uniform Distributed Load anywhere on                     '''
    '''     a left side cantilever                                                 '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at left Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim c As Double     'Width of load area
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim c5 As Double    'Integration constant - C5
    Dim c6 As Double    'Integration constant - C6
    Dim w_tot As Double 'Area of load or equivalent point load
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'left End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If w = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        cant_left_uniform_load = 0
    Else
        
        If a > L Or b > L Or a = b Or a < 0 Or b < 0 Then
            cant_left_uniform_load = "A,B cannot be equal and must be between 0 and L"
        ElseIf b < a Then
            cant_left_uniform_load = "B must be greater than A"
        Else
            ''''''''''''''''''''''''''''
            ''  Common Calculations   ''
            ''''''''''''''''''''''''''''
            
            c = b - a
            w_tot = w * c
            
            'Support Reactions
            rl = 0
            rr = w_tot
            mr = -1 * w_tot * (L - (a + (c / 2)))
            
            'Integration Constants
            
            c5 = 0 + (0.5 * w_tot * (L - (a + (0.5 * c))) * (L - (a + (0.5 * c))))
            c6 = ((1 / 6) * w_tot * (L - (a + (0.5 * c))) * (L - (a + (0.5 * c))) * (L - (a + (0.5 * c)))) - (c5 * L)
            c3 = ((-0.5) * w_tot * (b - (a + (0.5 * c))) * (b - (a + (0.5 * c)))) + c5 + ((1 / 6) * w * (b - a) * (b - a) * (b - a))
            c1 = c3
            c4 = ((-1 / 6) * w_tot * (b - (a + (0.5 * c))) * (b - (a + (0.5 * c))) * (b - (a + (0.5 * c)))) + (c5 * b) + c6 + ((1 / 24) * w * (b - a) * (b - a) * (b - a) * (b - a)) - (c3 * b)
            c2 = (c3 * a) + c4 - (c1 * a)
        
            
            '''''''''''''''''''''''''''
            ''  Result Selection     ''
            '''''''''''''''''''''''''''
            
            If result = 0 Then
                
                'Left Support Reaction
                cant_left_uniform_load = rl
                
            ElseIf result = 1 Then
                
                'left Support Reaction
                cant_left_uniform_load = rr
            
            ElseIf result = 2 Then
                
                'Shear at x
                If 0 <= x And x <= a Then
                    vx = 0
                ElseIf a < x And x <= b Then
                    vx = -1 * (w * (x - a))
                ElseIf b < x And x <= L Then
                    vx = -1 * w_tot
                Else
                    vx = 0
                End If
          
                cant_left_uniform_load = vx
            
            ElseIf result = 3 Then
            
                'Moment at x
                If 0 <= x And x <= a Then
                    mx = 0
                ElseIf a < x And x <= b Then
                    mx = -0.5 * w * (x - a) * (x - a)
                ElseIf b < x And x <= L Then
                    mx = -1 * w_tot * (x - (a + (0.5 * c)))
                Else
                    mx = 0
                End If
                
                cant_left_uniform_load = mx
                    
            ElseIf result = 4 Then
                
                'Cross Section Rotation/Slope at x
                If 0 <= x And x <= a Then
                    sx = c1 / (E * I)
                ElseIf a < x And x <= b Then
                    sx = (((-1 / 6) * w * (x - a) * (x - a) * (x - a)) + c3) / (E * I)
                ElseIf b < x And x <= L Then
                    sx = ((-0.5 * w_tot * (x - (a + (0.5 * c))) * (x - (a + (0.5 * c)))) + c5) / (E * I)
                Else
                    sx = 0
                End If
                
                cant_left_uniform_load = sx
            
            ElseIf result = 5 Then
                
                'Deflection at x
                If 0 <= x And x <= a Then
                    dx = ((c1 * x) + c2) / (E * I)
                ElseIf a < x And x <= b Then
                    dx = (((-1 / 24) * w * (x - a) * (x - a) * (x - a) * (x - a)) + c3 * x + c4) / (E * I)
                ElseIf b < x And x <= L Then
                    dx = (((-1 / 6) * w_tot * (x - (a + (0.5 * c))) * (x - (a + (0.5 * c))) * (x - (a + (0.5 * c)))) + (c5 * x) + c6) / (E * I)
                Else
                    dx = 0
                End If
                
                cant_left_uniform_load = dx
            
            ElseIf result = 6 Then
                'Fixed End Moment Left
                femL = 0
                
                cant_left_uniform_load = femL
            
            ElseIf result = 7 Then
            
                'Fixed End Moment left
                femR = mr
                
                cant_left_uniform_load = femR
            
            Else
            
                cant_left_uniform_load = CVErr(xlErrNA)
                    
            End If
        End If
    End If
    
End Function

Function cant_left_variable_load(w1 As Double, w2 As Double, a As Double, b As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Variable Distributed Load anywhere on                    '''
    '''     a left side cantilever                                                 '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at left Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim c As Double     'Width of load area
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim c5 As Double    'Integration constant - C5
    Dim c6 As Double    'Integration constant - C6
    Dim c7 As Double    'Integration constant - C7
    Dim w As Double     'Area of load or equivalent point load
    Dim s As Double     'Slope of load area
    Dim dl As Double    'Center of load area from left
    Dim dr As Double    'Center of load area from right
    Dim cc As Double
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'left End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If w1 = 0 And w2 = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        cant_left_variable_load = 0
    Else
    
        If a > L Or b > L Or a = b Or a < 0 Or b < 0 Then
            cant_left_variable_load = "A,B cannot be equal and must be between 0 and L"
        ElseIf b < a Then
            cant_left_variable_load = "B must be greater than A"
        ElseIf Sgn(w1) <> Sgn(w2) And w1 <> 0 And w2 <> 0 Then
            cant_left_variable_load = "W1,W2 must have the same sign"
        Else
            ''''''''''''''''''''''''''''
            ''  Common Calculations   ''
            ''''''''''''''''''''''''''''
            
            c = b - a
            w = 0.5 * (w1 + w2) * c
            dl = a + (((w1 + (2 * w2)) / (3 * (w2 + w1))) * c)
            dr = L - dl
            s = (w1 - w2) / c
            cc = (((w1 + (2 * w2)) / (3 * (w2 + w1))) * c) + a
            
            'Support Reactions
            rl = 0
            rr = w
            mr = -1 * rr * (L - cc)
            
            'Integration Constants
            c6 = 0 + (0.5 * w * (L - cc) * (L - cc))
            c7 = ((1 / 6) * w * (L - cc) * (L - cc) * (L - cc)) - (c6 * L)
            c3 = -1 * ((1 / 6) * a * ((a * a * s) - (3 * a * ((a * s) + w1)) + (3 * a * ((a * s) + (2 * w1)))))
            c4 = (-0.5 * w * (b - cc) * (b - cc)) + c6 - (c3 * b) - ((1 / 24) * b * b * ((b * b * s) - (4 * b * ((a * s) + w1)) + (6 * a * ((a * s) + (2 * w1)))))
            c5 = ((-1 / 6) * w * (b - cc) * (b - cc) * (b - cc)) + (c6 * b) + c7 - (0.5 * c3 * b * b) - (c4 * b) - ((1 / 120) * b * b * b * ((b * b * s) - (5 * b * ((a * s) + w1)) + (10 * a * ((a * s) + (2 * w1)))))
            c1 = ((1 / 24) * a * a * ((a * a * s) - (4 * a * ((a * s) + w1)) + (6 * a * ((a * s) + (2 * w1))))) + (c3 * a) + c4
            c2 = ((1 / 120) * a * a * a * ((a * a * s) - (5 * a * ((a * s) + w1)) + (10 * a * ((a * s) + (2 * w1))))) + (0.5 * c3 * a * a) + (c4 * a) + c5 - (c1 * a)
            
            '''''''''''''''''''''''''''
            ''  Result Selection     ''
            '''''''''''''''''''''''''''
            
            If result = 0 Then
                
                'Left Support Reaction
                cant_left_variable_load = rl
                
            ElseIf result = 1 Then
                
                'left Support Reaction
                cant_left_variable_load = rr
            
            ElseIf result = 2 Then
                
                'Shear at x
                If 0 <= x And x <= a Then
                    vx = 0
                ElseIf a < x And x <= b Then
                    vx = (-0.5 * ((2 * w1) - (s * (x - a)))) * (x - a)
                ElseIf b < x And x <= L Then
                    vx = -1 * rr
                Else
                    vx = 0
                End If
            
                cant_left_variable_load = vx
            
            ElseIf result = 3 Then
            
                'Moment at x
                If 0 <= x And x <= a Then
                    mx = 0
                ElseIf a < x And x <= b Then
                    mx = ((1 / 6) * x * ((x * x * s) - (3 * x * ((a * s) + w1)) + (3 * a * ((a * s) + (2 * w1))))) + c3
                ElseIf b < x And x <= L Then
                    mx = -1 * w * (x - cc)
                Else
                    mx = 0
                End If
                
                cant_left_variable_load = mx
                    
            ElseIf result = 4 Then
                
                'Cross Section Rotation/Slope at x
                If 0 <= x And x <= a Then
                    sx = c1 / (E * I)
                ElseIf a < x And x <= b Then
                    sx = (((1 / 24) * x * x * ((x * x * s) - (4 * x * ((a * s) + w1)) + (6 * a * ((a * s) + (2 * w1))))) + (c3 * x) + c4) / (E * I)
                ElseIf b < x And x <= L Then
                    sx = ((-0.5 * w * (x - cc) * (x - cc)) + c6) / (E * I)
                Else
                    sx = 0
                End If
                
                cant_left_variable_load = sx
            
            ElseIf result = 5 Then
                
                'Deflection at x
                If 0 <= x And x <= a Then
                    dx = ((c1 * x) + c2) / (E * I)
                ElseIf a < x And x <= b Then
                    dx = (((1 / 120) * x * x * x * ((x * x * s) - (5 * x * ((a * s) + w1)) + (10 * a * ((a * s) + (2 * w1))))) + (0.5 * c3 * x * x) + (c4 * x) + c5) / (E * I)
                ElseIf b < x And x <= L Then
                    dx = (((-1 / 6) * w * (x - cc) * (x - cc) * (x - cc)) + (c6 * x) + c7) / (E * I)
                Else
                    dx = 0
                End If
                
                cant_left_variable_load = dx
            
            ElseIf result = 6 Then
                'Fixed End Moment Left
                femL = 0
                
                cant_left_variable_load = femL
            
            ElseIf result = 7 Then
            
                'Fixed End Moment left
                femR = mr
                
                cant_left_variable_load = femR
            
            Else
            
                cant_left_variable_load = CVErr(xlErrNA)
                    
            End If
        End If
    End If
End Function


